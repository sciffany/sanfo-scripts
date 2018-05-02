var infoArray = ['1a4PgQxC8wO2vE96e5or0Jl4mIOJciBbxupZ1Z_zgDpc',
                 '1ej-drAEs_MPLwY8YveJOUOqzz6GA5IdjJKFbX1uKCkQ',
                 '1-PD7vyJE0hRLJ-T0QVkewOXDZBTMGTWkctRRXXuuPvs'];
var columns = [1,13]; //columns to check the date for dscohsheets
var outputColumnStart = 2; //on which column will we paste the results
var DIFF = 9; //how far COH is from the DATE
var cohSheetName = "Sheet1";
var dscohSheetName = "DSCOH";

var today = new Date();
var cohSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cohSheetName);

//what it does daily
function updateDaily(){
  var dateColumnData = cohSheet.getRange(1,1,cohSheet.getMaxRows()).getValues();
  var outputRow=__searchDate__(dateColumnData, today);
  var outputColumn = outputColumnStart;
  for (var i=0; i<infoArray.length; i++){
    for (var j=0; j<columns.length; j++){
      __autoUpdate__(i, columns[j], outputColumn, outputRow, cohSheet);
      outputColumn++;
    }
  }
}

//what it does for 6 files
function __autoUpdate__(stnNo, sourceColumn, outputColumn, outputRow){
  var dscohSheet = SpreadsheetApp.openById(infoArray[stnNo]).getSheetByName(dscohSheetName);
  var dateColumnData = dscohSheet.getRange(1,sourceColumn,dscohSheet.getMaxRows()).getValues();
  var sourceRow = __searchDate__(dateColumnData, today);
  var ans = dscohSheet.getRange(sourceRow+1, sourceColumn+DIFF).getValue();
  cohSheet.getRange(outputRow+1, outputColumn,1,1).setValue(ans);
}

//given a matrix and a date, find the correct row
function __searchDate__(dateColumnData, startDate){
  for (var i=0;i<dateColumnData.length;i++) {
    var thisDate = dateColumnData[i][0];
    if (thisDate instanceof Date &&
        startDate.getDate() == thisDate.getDate() &&
        startDate.getMonth() == thisDate.getMonth() &&
        startDate.getYear() == thisDate.getYear()) {
          return i;
    }
  }
}