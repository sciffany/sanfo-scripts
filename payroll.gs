function autoCount()
{
  var DAYCLASSROW = 15;
  var now = new Date();
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Please type in starting month number' , ui.ButtonSet.OK_CANCEL);
  var monthNo = parseInt(response.getResponseText());
  //if (monthNo/10<1){monthNo="0"+monthNo;}
  var yearNo = now.getYear();
  
  if (response.getSelectedButton()=='CANCEL'){return 0;}
  if (monthNo==12)
  {  
    var response2 = ui.prompt('Since you are updating December payroll, type in the year', ui.ButtonSet.OK_CANCEL);
    yearNo = parseInt(response2.getResponseText());
    if (response2.getSelectedButton()=='CANCEL'){return 0;}
  
  }
  

  {     
    var compSheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Computation")
    var selectedCell = compSheet.getActiveRange();
    noOfDays = (selectedCell.getColumn() - 2)/2;
    noOfEmployees = selectedCell.getRow() - 4;
    
    //storing all the important info
    var range = compSheet.getRange(5,3, noOfEmployees, noOfDays*2);
    var values = range.getValues();
   
    //getting the first day
    var firstDate = compSheet.getRange(3,3).getValues()[0][0];
    var firstDay = new Date(yearNo, monthNo-1, firstDate);
    
    //creating an array to store types of day
    var dateType = new Array(noOfDays+1).join('0').split('').map(parseFloat);
    var dateArray = ['r', 's', 'rh', 'sh'];
    
    //mark all sundays   
    var firstSun = (7-firstDay.getDay())%7;
    while(firstSun<noOfDays)
    {
      dateType[firstSun]=1;
      firstSun+=7;
    }
    
    //parsing the holiday sheet
    var holidaySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Holidays");
    var lastRow = holidaySheet.getLastRow();
    var holidayList = holidaySheet.getRange(2,1,lastRow-1,3); //getting all the regular and special holidays
    var hValues = holidayList.getValues();   
    var holidayColumn = [0,2];
    for (var i=0; i<2; i++)
    {
      for (var j=0; j<lastRow-1; j++)
      {
        var currHol = hValues[j][holidayColumn[i]];
        
        
        if (currHol==''){continue;}
        var diff = (currHol - firstDay)/(24*60*60*1000);
        diff%=365;
        
        if(diff>=0 && diff<noOfDays)
        {
          dateType[diff]=i+2;
        }
        
      }
      
    }
    
    //marking of holidays
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(compSheet);
    for (var i=0;i<noOfDays;i++)
    {
      compSheet.getRange(DAYCLASSROW,i*2+3).setValue(dateArray[dateType[i]]);
      compSheet.getRange(DAYCLASSROW,i*2+4).setValue(dateArray[dateType[i]]);
    }
    
    //analysing every employee
    for (var i=0;i<noOfEmployees;i++)
    {
      var employeeLine = values[i];
      var empAttendance = new Array(9+1).join('0').split('').map(parseFloat);
      var daysPresent = 0;
      for (var j=0; j<noOfDays;j++)
      {
        //regular
        if (employeeLine[j*2]!='')
        {
          empAttendance[dateType[j]*2] += employeeLine[j*2];
          if (dateType[j]!=2) //not regular holiday, and you go to work
          {
            daysPresent+=1
          }
        }
        
        //OT
        if (employeeLine[j*2+1]!='')
        {empAttendance[dateType[j]*2+1] += employeeLine[j*2+1];}
       
      
      }  
      //updating daysPresent
      empAttendance[9] = daysPresent;
      
      //pasting his attendance
      for (var k=0;k<empAttendance.length;k++)
      {
        compSheet.getRange(i+5,40+k).setValue(empAttendance[k]);
      }
    }

  }
  
}
