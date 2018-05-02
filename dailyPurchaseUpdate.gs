var letter1 = 'b'
var letter2 = 'm'
var name1 = 'Bicutan'
var name2 = 'Manda'
var DSCOHsheet = "DSCOH"

function updateBicMan()
{
  updateBicutan();
  updateManda();
}

function updateBicutan()
{
 updateDaily_(letter1, name1, 1,0);
}


function updateManda()
{
 updateDaily_(letter2, name2, 2,0);
}

function printBicutan()
{
 printWeekly_(letter1, name1);
}

function printManda()
{
 printWeekly_(letter2, name2);
}



function updateDaily_(letter, branch,pageloc,mode)
{
  
  var activeRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DSCOHsheet).getActiveRange();
  activeRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DSCOHsheet).getRange(activeRange.getRow()-1, activeRange.getColumn(), activeRange.getHeight()+1, activeRange.getWidth()+1);
  var dValues = activeRange.getValues();
 
  var startDate = dValues[1][0].getDate();
  var endDate = dValues[dValues.length-1][0].getDate();
  var emailStr;

  if (mode)
  {
    var emailStr="Updated " + letter + startDate + " to " + letter + endDate;
  }
  else{
    var warning = 'Will update from dates ' + letter + startDate + " to " + letter + endDate + ". Please ensure that all the dates fall within the range of the excel file.";
    var ui = SpreadsheetApp.getUi();
    var response= ui.alert('Daily Update for ' + branch, warning, ui.ButtonSet.OK_CANCEL);
  }
  
  if (response == 'CANCEL')
  {return 0;}
  for (var i = 1; i < dValues.length; i++)
  {
    
    doingDate = dValues[i][0].getDate();
    firstCN = dValues[i-1][1]+1;
    lastCN = dValues[i][1];
    noOfEntries = lastCN - firstCN +1;
    
    
    
    if (noOfEntries>0)
      
    {
      var sheetName = "JC" + letter.toUpperCase();
      var JCsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      var range = JCsheet.getRange(7,3,20000,1)
      var values = range.getValues();
      
      var firstRow = 0;
      for (var j = 0; j < values.length; j++)
      {
        if (values[j][0] == firstCN)
        {
          var today = new Date();
          var thisYear = today.getYear();
          if (JCsheet.getRange(j+7,1).getValue().getYear()==thisYear)
              {
            firstRow = j+7;
            break;
            }
        }
      }
      
      var dataRange = JCsheet.getRange(firstRow, 1, noOfEntries, 10);
      var dataValues = dataRange.getValues();
      
      var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(letter+doingDate);
      destinationSheet.getRange(7,1, 200, 10).clearContent();
      destinationSheet.getRange(7, 1,noOfEntries,10).setValues(dataValues);
    }
      
  }
  
    if (!mode)
    {
      ui.alert('Now comparing ' +branch+ ' with summary');
    }
    var xRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DSCOH").getRange(activeRange.getRow()+1, activeRange.getColumn()+2, activeRange.getHeight()-1, activeRange.getWidth()+5);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DSCOH").setActiveRange(xRange);
    var xValues = xRange.getValues();
    
    var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
    var lRange = summarySheet.getRange(11*pageloc-8, 1, 7,1);
    var lValues = lRange.getValues();
    
    var firstRow = 0;
    for (var j = 0; j < lValues.length; j++)
    {
      if (lValues[j][0] == startDate)
      {
        firstRow = 11*pageloc-8+j;
        break;
      }
    }
 
    
    var yRange = summarySheet.getRange(firstRow, 2, activeRange.getHeight()-1,7);
    summarySheet.setActiveRange(yRange);
    var yValues = yRange.getValues();
    
    var mistakes= false;
    for(var i=0; i < xValues.length ; i++)
    {
      for(var j=0; j <xValues[0].length;j++)
      {          
        
        if(xValues[i][j]!=Math.round(yValues[i][j]*100)/100)
        {
          var wrongRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DSCOH").getRange(xRange.getRow()+i, xRange.getColumn()+j, 1,1);
          wrongRange.setBackground("red");
          mistakes=true;
        }
      }
    }
    
    
    
    summarySheet.getRange(1, 7).setValues([[dValues[dValues.length-1][0]]]);
    
    
    if (mistakes)
    {
      if(mode)
      {
        emailStr+="\nMistakes found!";
      }  
      else
      {  
        var alert = ui.alert('Mistakes found!', ui.ButtonSet.OK);
      }
    }
    else
    { 
      if(mode)
      {
        emailStr+="\nPerfect!";
      }  
      else
      {  
        var alert = ui.alert('Perfect!', ui.ButtonSet.OK);
      }
    }
    
    if (mode)
    {
      var email = "bills.to.rfp@hotmail.com";
      var email2 = "sciffany@gmail.com";
      var subject = "UDPATE "+branch;
      GmailApp.sendEmail(email, subject, emailStr);
      GmailApp.sendEmail(email2, subject, emailStr);
    }
    
    if (pageloc==1)
    {
      var zRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DSCOH").getRange(activeRange.getRow()+1, activeRange.getColumn()+12, activeRange.getHeight()-1, activeRange.getWidth()-1);
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DSCOH").setActiveRange(zRange);
      var zValues = zRange.getValues();
    }
    
    
  
  
}



function clearWeekly()
{
  
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Clearing Weekly Sheets', 'What is the new starting date?', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton()=='OK')
    
  {
    var startDate = parseInt(response.getResponseText());
    var endDate = startDate + 6
    
    //get new dates
    var dates = new Array();
    if(startDate>=22)
    {
      var response = ui.prompt('Possible end of the month', 'What day does this month end in?', ui.ButtonSet.OK_CANCEL);
      var endMonth = parseInt(response.getResponseText());
      
      if (endMonth < endDate)
      {
        for(var i = startDate; i <= endMonth; i++)
        {
          dates.push(i);
        }
        
        for(var j = 1; j < endDate-endMonth+1; j++)
        {
          dates.push(j);
        }
      }
      
      else
      {
        for (var i = startDate; i <= endDate; i++)
        {
          dates.push(i);
        }
      }
      
    }
    
    else
    {
      for (var i = startDate; i <= endDate; i++)
      {
        dates.push(i);
      }
    }
    
    
    /**clears and renames sheets*/
    var b=0;
    var m=0;
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i=0; i<sheets.length ; i++)
      
    {
      
      if (sheets[i].getName()[0] == letter1)
        
      {
        sheet = sheets[i];
        sheet.getRange(7,1, 200, 10).clearContent();
        sheet.setName(letter1+dates[b]);
        b++;
      }
      
      else if (sheets[i].getName()[0] == letter2)
        
      {
        sheet = sheets[i];
        sheet.getRange(7,1, 200, 10).clearContent();
        sheet.setName(letter2+dates[m]);
        m++;
      }
      
    }
    
   /** resets Summary sheet dates*/
    var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
    for (var i=0;i<7;i++)
    {
      summarySheet.getRange(i+3,1).setValue(dates[i]);
      summarySheet.getRange(i+14,1).setValue(dates[i]);
    }

    
    var alert = ui.alert('All done!', ui.ButtonSet.OK);
    
    
    
  }
  
}


function printWeekly_(letter, branch)
{
  var ui = SpreadsheetApp.getUi();
  var firstResponse= ui.prompt('Weekly print for ' + branch, 'What is the first control number?', ui.ButtonSet.OK_CANCEL);
  var firstCN = parseInt(firstResponse.getResponseText());
  
  var lastResponse = ui.prompt('Weekly for ' + branch, 'What is the last control number?', ui.ButtonSet.OK_CANCEL);
  var lastCN = parseInt(lastResponse.getResponseText());
  var noOfEntries = lastCN - firstCN+1;
  
  if (lastResponse.getSelectedButton() =='OK' && firstResponse.getSelectedButton() == 'OK')
  {
    
    var sheetName = "JC" + letter.toUpperCase();
    var JCsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var range = JCsheet.getRange(7,3,5570,1)
    var values = range.getValues();
    
    var firstRow = 0;
    for (var i = 0; i < values.length; i++)
    {
      if (values[i][0] == firstCN)
      {
        var today = new Date();
          var thisYear = today.getYear();
          if (JCsheet.getRange(i+7,1).getValue().getYear()==thisYear)
              {
            firstRow = i+7;
            break;
            }
      }
    }
    
    var dataRange = JCsheet.getRange(firstRow, 1, noOfEntries, 10);
    var dataValues = dataRange.getValues();
    
    var destinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Print"+letter.toUpperCase());
    SpreadsheetApp.setActiveSheet(destinationSheet);
    destinationSheet.getRange(7,1, 1200, 10).clearContent();
    destinationSheet.getRange(7, 1, noOfEntries,10).setValues(dataValues);
    
  }
}


function autoUpdate()
{
  var today = new Date();
  var day = today.getDay();
  var startDate = today;
  var noOfDates=1;
  if (day==1 || day==2)
  {return 0;}
  else if(day==3)//wednesday
  {
    clearWed_();
    startDate.setDate(startDate.getDate()-3);
    noOfDates =3;
  }
  else
  {
    startDate.setDate(startDate.getDate()-1);
  }
  startDate = new Date(startDate.getYear(),startDate.getMonth(),startDate.getDate());
  var startIndex=0;
  var dscohSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DSCOH");
  for (var i=1;i<dscohSheet.getMaxRows();i++)
  {
    var thisDate = dscohSheet.getRange(i,1).getValue();
    if (!(thisDate instanceof Date))
    {continue;}
    if (startDate.getDate() == thisDate.getDate())
    {
      if (startDate.getMonth() == thisDate.getMonth())
      {
        if (startDate.getYear() == thisDate.getYear())
        {
          startIndex =i;
          break;
        }
      }
    }
  }
  dscohSheet.setActiveRange(dscohSheet.getRange(startIndex,1,noOfDates,1));
  
  updateDaily_(letter1, name1, 1, 1);
  updateDaily_(letter2, name2, 2, 1);
  
}


function clearWed_()
{
  var dates = new Array();
  var today = new Date();
  
  //sunday
  today.setDate(today.getDate()-3);
  dates.push(today.getDate());
  
  //rest of the dates
  for (var i=1; i<7;i++)
  {
    today.setDate(today.getDate()+1);
    dates.push(today.getDate());
  
  }
  
  /**clears and renames sheets*/
  var b=0;
  var m=0;
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0; i<sheets.length ; i++)
    
  {
    
    if (sheets[i].getName()[0] == letter1)
      
    {
      sheet = sheets[i];
      sheet.getRange(7,1, 200, 10).clearContent();
      sheet.setName(letter1+dates[b]);
      b++;
    }
    
    else if (sheets[i].getName()[0] == letter2)
      
    {
      sheet = sheets[i];
      sheet.getRange(7,1, 200, 10).clearContent();
      sheet.setName(letter2+dates[m]);
      m++;
    }
    
  }
  
  /** resets Summary sheet dates*/
  var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
  for (var i=0;i<7;i++)
  {
    summarySheet.getRange(i+3,1).setValue(dates[i]);
    summarySheet.getRange(i+14,1).setValue(dates[i]);
  }
  
  
}




