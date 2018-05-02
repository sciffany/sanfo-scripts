var PAYEE_COLUMN = 9;
var CHECK_NO_COLUMN = 3;
var DATE_COLUMN = 'B';
var DATE_COLUMN_NO = 2;
var AMOUNT_COLUMN = 10;
var NO_OF_CONTENT_COLUMNS = 11;

// Convert numbers to words
// copyright 25th July 2006, by Stephen Chapman http://javascript.about.com
// permission to use this Javascript on your web page is granted
// provided that all of the code (including this copyright notice) is
// used exactly as shown (you can change the numbering system if you wish)

// American Numbering System
var th = ['','thousand','million', 'billion','trillion'];

var dg = ['zero','one','two','three','four',
'five','six','seven','eight','nine']; var tn =
['ten','eleven','twelve','thirteen', 'fourteen','fifteen','sixteen',
'seventeen','eighteen','nineteen']; var tw = ['twenty','thirty','forty','fifty',
'sixty','seventy','eighty','ninety'];

function toWords_(s)
{
s = s.toString();
s = s.replace(/[\, ]/g,'');
if (s != parseFloat(s)) return 'not a number';
var x = s.indexOf('.'); if (x == -1) x = s.length; if (x > 15) return 'too big';
var n = s.split('');
var str = '';
var sk = 0;
for (var i=0; i < x; i++)
{
  if ((x-i)%3==2) {if (n[i] == '1')
  {str += tn[Number(n[i+1])] + ' '; i++; sk=1;}
  else if (n[i]!=0)
  {str += tw[n[i]-2] + ' ';sk=1;}} 
  else if (n[i]!=0)
  {str += dg[n[i]] +' '; if ((x-i)%3==0) str += 'hundred ';sk=1;}
  if ((x-i)%3==1)
  {if (sk)
  str += th[(x-i-1)/3] + ' ';sk=0;}
}

return str.replace(/\s+/g,' ');
  
  
}

function toMoney_(m)
{
  var ans="* ";
  ans+= toWords_(parseInt(m));
  
  if (parseInt((m*100)%100)!=0){ //if there are cents
    ans+="and ";
    ans+= parseInt((m*100)%100);
    ans+="/100 ";
  }
  ans+="*";
  return ans;
   
}

var infobmm = ['1R1RQx9QsM7vF7grrEHsh4N509NE1gNMyrKNipPs34hc', 'BICUTAN', '1SLSquYebno8i_jDoUBzpsq5Suaacpa4zY_4xptmNSfY', 'MANDA', '1YZCuiVMTFpc7l-4LYky9TAZ_HCiNjnY8m3A1Sgu7paI', 'MARIQ'];
var infoivk = ['1huYkG4yUAhFxE-l-4qTqLdO44XvX6FoE8kRPVNdGt8Q', 'IMUS', '1CDtdp5A3CVGHKBqj5JhVswQR-ks5VROVkzoEETvXA2k', 'KAINGIN', '1hJb7u3sqGNC_nR_GoaP2oPEdhnK85yCml8zzfGn9dSk', 'VAL'];
var infoc = ['1KJv7-Mg4vj1MDQxk3QLuIvXKxdBblKIaqMM-th-rxMU', 'MSC Central'];
var infow = ['1oTW6jXoQiVWtvYnxMWL-qyFXBt-jj4tlhqSg3i2wpLg', 'Weekly'];
var colbmm =['#4a86e8','#ff00ff','#85200c'];
var colivk=['#3b2eae','#bf9000','#9900ff'];
var colc=['#00ffff'];
var colw=['#b7b7b7'];

//for the check file
var checkKey = '1vfzdJ8AJXHS4kFKi7PcG03jnKiFTWGYEs4RRno6tq00';
var checkSS = SpreadsheetApp.openById(checkKey);
//for the voucher file
var voucherKey = '1LfuQu8u0kryC7t614JAtdabPeT24IGesn81VDQgJzSM';
var voucherSS = SpreadsheetApp.openById(voucherKey);
//for this own file
var thisSS = SpreadsheetApp.getActiveSpreadsheet();
//sets voucher and check sheets
var thisSheet = thisSS.getSheetByName("BICUTAN");
var voucherSheet = thisSS.getSheetByName("Voucher");
var checkSheet = thisSS.getSheetByName("Check");


function bicManMarRFP(){
  
  toCV_(infobmm, colbmm);
  
}


function ImValKaiRFP(){
  
  toCV_(infoivk,colivk);
  
}


function centralRFP(){
  toCV_(infoc,colc);
  
}


function weeklyRFP(){
  
  toCV_(infow,colw);
  
}

function clearSheets()
{
  clearSheets_(voucherSS);
  clearSheets_(checkSS);
}  

function clearSheets_(ss)
{
  ss.insertSheet(0);
  ss.getSheets();
  var sheets = ss.getSheets();
  
  for (i = 1; i < sheets.length; i++)
  { 
    ss.deleteSheet(sheets[i]);
  }
}

function toPdf_(ss)
{
  var blob = DriveApp.getFileById(ss.getId()).getAs("application/pdf");
  blob.setName(ss.getName() + ".pdf");
  return blob;
}

function toCV_(infoArray, color)
{
  
  var vSheetCounter = voucherSS.getSheets().length;
  var cSheetCounter = checkSS.getSheets().length;

  var index=0;
  var ui=SpreadsheetApp.getUi();
  
  while (index<infoArray.length)
  {
    //opens shared file
    var fileKey = infoArray[index];
    index++;
    var fileSS = SpreadsheetApp.openById(fileKey);
    var fileSheet = fileSS.getSheetByName(infoArray[index]);
    
    var column = DATE_COLUMN;
    var lastRow = fileSheet.getMaxRows();
    var values = fileSheet.getRange(column + "1:" + column + lastRow).getValues();
    for (; values[lastRow - 1] == "" && lastRow > 0; lastRow--){}
    
    var date = fileSheet.getRange(lastRow, DATE_COLUMN_NO).getValue();
    var green='#00ff00';          
    ui = SpreadsheetApp.getUi();
    
    //get lastUpdate
    var lastUpdate = lastRow-1;
    var columnOfDates = fileSheet.getRange(1, DATE_COLUMN_NO, lastRow).getValues();
    
    for (; columnOfDates[lastUpdate] == date && lastUpdate > 0; lastUpdate--){}
    lastUpdate+=2;
    
    //get lastRow
    var lastRow = lastUpdate;
    for (; fileSheet.getRange(lastRow, CHECK_NO_COLUMN).getBackground() != green; lastRow++){};
    
    //change colour of last row
    fileSheet.getRange(lastRow, CHECK_NO_COLUMN,1,1).setBackground('#d9ead3');
    
    //changes check voucher title
    var titlePos = 'C1';
    var titlePos2 = 'C23';
    voucherSheet.getRange(titlePos).setValue(infoArray[index] + " Check Voucher");
    voucherSheet.getRange(titlePos).setFontColor(color[(index+1)/2-1]);
    voucherSheet.getRange(titlePos2).setFontColor(color[(index+1)/2-1]);
    
    for (var i=lastUpdate;i<=lastRow;i++)
    {
      var amountInWords = toMoney_(fileSheet.getRange(i, AMOUNT_COLUMN).getValue());
      //copies line
      values = fileSheet.getRange(i, 1, 1, NO_OF_CONTENT_COLUMNS).getValues();
      //pastes it to above
      thisSheet.getRange(3, 1, 1, NO_OF_CONTENT_COLUMNS).setValues(values);
      thisSheet.getRange(3, NO_OF_CONTENT_COLUMNS+1).setValue(amountInWords);
      
      //copies sheet contents of voucher
      var voucherV = voucherSheet.getRange(1, 1, voucherSheet.getMaxRows(), voucherSheet.getMaxColumns()).getDisplayValues();
      
      //gets voucher sheet to reqforpay file
      voucherSheet.copyTo(voucherSS);
      var vSheet = voucherSS.getSheets()[vSheetCounter];
      
      vSheetCounter++;
      
      //pastes voucher contents
      vSheet.getRange(1,1,voucherSheet.getMaxRows(), voucherSheet.getMaxColumns()).setValues(voucherV);
      
      //gets payeeName
      var payeeName = thisSheet.getRange(3, PAYEE_COLUMN).getDisplayValue() + " ("+ infoArray[index] +")";
      
      //if check doesn't exist
      var currCheck = checkSS.getSheetByName(payeeName);
      if (currCheck==null)
      {
        //copies sheet contents of check
        var checkV = checkSheet.getRange(1, 1, checkSheet.getMaxRows(), checkSheet.getMaxColumns()).getDisplayValues();
        //gets check sheet to reqforpay file
        checkSheet.copyTo(checkSS);
        var cSheet = checkSS.getSheets()[cSheetCounter];
        cSheetCounter++;
        cSheet.setName(payeeName);
        //pastes check contents
        cSheet.getRange(1,1,checkSheet.getMaxRows(), checkSheet.getMaxColumns()).setValues(checkV);
      }
      
      else
      {
        var newAmt = currCheck.getRange(5,8).getValue() + checkSheet.getRange(5,8).getValue();
        currCheck.getRange(5,8).setValue(newAmt);
        currCheck.getRange(7,2).setValue(toMoney_(newAmt));
      }  
    }
    index++;
  }
  
}

function compileToPDF()
{
  //deletes the first page
  voucherSS.deleteSheet(voucherSS.getSheets()[0]);
  checkSS.deleteSheet(checkSS.getSheets()[0]);
  
  var blob = toPdf_(voucherSS);
  
  /*Source of PDF converter
  https://ctrlq.org/code/19869-email-google-spreadsheets-pdf
  */
  
  // Send the PDF of the spreadsheet to this email address
  var email = "sciffany@gmail.com";
  
  // Subject of email message
  var subject = "Voucher PDF generated from spreadsheet" 

  // If allowed to send emails, send the email with the PDF attachment
  if (MailApp.getRemainingDailyQuota() > 0) 
    GmailApp.sendEmail(email, subject, "Pdf is attached. View your check at https://docs.google.com/spreadsheets/d/1vfzdJ8AJXHS4kFKi7PcG03jnKiFTWGYEs4RRno6tq00/edit#gid=1586562825", {
      attachments:[blob]
    });
}


