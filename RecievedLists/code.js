
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Check sheets for DB')
        .addItem('Select Date', 'createDialog')
        .addToUi();
        
}
function onEdit(e){
    var range = e.range;
    var sheet = SpreadsheetApp.getActiveSheet();
    if (range.getColumn() == 8 && range.getValue() == "done")
    {
        sheet.getRange(range.getRow(), range.getColumn()).setValue("Done");
    }
}
function createDialog() {

    var htmlDialog = HtmlService.createHtmlOutputFromFile("scriptHTML")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(200)
        .setWidth(270);
    SpreadsheetApp.getUi().showModalDialog(htmlDialog, "Select Date");
}
function getMissingColumns(sheet, columnNames, newColumnNames)
{
try{
    var range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    var values = range.getValues();

    for ( i in values[0] )
        values[0][i] = replaceSpaces(values[0][i])
    for (column in columnNames)
    {
         var index = values[0].indexOf( columnNames[column]); //UPDATE 28.11   !!!!!!!!!!!!!!!!!!!!!!!
         if (index!= -1 ) {
             newColumnNames[columnNames[column]] = index;
             }
         else newColumnNames[columnNames[column]] = columnNames[column];
       
     }   
   
     
     
      var mistakes_test = "";
     for (column in columnNames)
     {
         if (typeof(newColumnNames[columnNames[column]] ) != 'number')
            mistakes_test+=newColumnNames[columnNames[column]] + "; "
     }
     range.setValues(values);

} catch (err) {Browser.msgBox(err)}

    return mistakes_test;
}
function cleanTheList(linkList) {

  var currentList;
  try {currentList   = SpreadsheetApp.openByUrl(linkList).getSheets()[0]; }
  catch (e) {return "no permission"; }

    var lastColumn = currentList.getLastColumn();
    var lastRow = currentList.getLastRow();
    var range = currentList.getRange(1, 1, lastRow, lastColumn);
    var values = range.getValues();
    var maxRows = currentList.getMaxRows();
    var countRows = maxRows - lastRow;
    
    if (countRows != 0) currentList.deleteRows(lastRow + 1, maxRows - lastRow); //delete empty rows;
    
    var columnNames = ["first_name","last_name","company","title","email","address","city","state","zip","country","phone","prooflink","employees","employees_prooflink","revenue","revenue_prooflink"];
    var newColumnNames=[] ;
    var mistakes =  getMissingColumns(currentList,columnNames,newColumnNames)
    
     var bgColors = range.getBackgrounds();
     var weights = range.getFontWeights();
     var fontColors = range.getFontColors();
     
   var newEmail = false;
   var newEmailRows ="" ;

    for (var i = 1; i < lastRow; i++) {
  
        var prooflink = newColumnNames['prooflink']

        if (typeof prooflink == 'number' ) {
            var link_str = values[i][prooflink].toString();
            if (link_str.indexOf('linkedin') != -1)
                values[i][prooflink] = link_str.split('?')[0];
        }
         var employees_prooflink = newColumnNames['employees_prooflink']

        if (typeof employees_prooflink == 'number') {
            link_str = values[i][employees_prooflink].toString();
            if (link_str.indexOf('yahoo') != -1 || link_str.indexOf('linkedin') != -1)
                values[i][employees_prooflink]= link_str.split('?')[0];
        }
        
        var revenue_prooflink =  newColumnNames['revenue_prooflink']
        if (typeof revenue_prooflink == 'number') {
            link_str = values[i][revenue_prooflink].toString();
            if (link_str.indexOf('yahoo') != -1 || link_str.indexOf('linkedin') != -1)
                values[i][revenue_prooflink] = link_str.split('?')[0];
        }
        var email = newColumnNames['email']
       if ( typeof email== 'number' && bgColors[0][email]!= "#f5bfb3" && bgColors[i][email] =="#ffff00" && fontColors[i][email]== "#ff0000" && weights[i][email]=="bold" )
           {
            currentList.getRange(1, email+1).setBackground("#f5bfb3") ;
            newEmailRows += (i+1) +", ";
            
        }
    }
    if (newEmailRows) { mistakes = "New emails found: " + newEmailRows + "missing " + mistakes; return mistakes}
    else if (mistakes!= ""){return "missing "+ mistakes; }

    var isUnCheckedList = true; //check if list is fully uncheked (no color coding for: title,phone, prooflink)
    
    for (var i = 1; i < lastRow; i++) {
        if (bgColors[i][newColumnNames['prooflink']] != "#ffffff") { isUnCheckedList = false; break; }

        else if (bgColors[i][newColumnNames['phone']]!= "#ffffff") { isUnCheckedList = false; break; }
        
        else if (bgColors[i][newColumnNames['title']] != "#ffffff") { isUnCheckedList = false; break; }
    }
    if (isUnCheckedList)
        return false;
    range.setValues(values);


    return true;
}

function useScript(dateToScript, DaySelected, MonthSelected, YearSelected, isWholeMonth, checkRejectionRate) {
  
  
    var currentDate= (generateCurrentDate());
    var recl = SpreadsheetApp.getActiveSpreadsheet();
    if (isWholeMonth) {
        var sheetToScript = SpreadsheetApp.getActiveSheet();
    }
    else {
        var nameOfSheet = getNameOfSheet(MonthSelected, YearSelected);
        var sheetToScript = recl.getSheetByName(nameOfSheet);
    }
    if (sheetToScript) {
   
        var lastColumn = sheetToScript.getLastColumn();
        
        if (!isWholeMonth){
            var dateRange = sheetToScript.getRange (1,1,sheetToScript.getLastRow());
            var dateValues = dateRange.getValues();
            var stopPosition =0;
            var startPosition=0;
       
            while (stopPosition < dateValues.length)
            {
               if (dateToScript ==dateValues[stopPosition] ) 
                 {
                     startPosition = stopPosition;
                     while (++stopPosition <dateValues.length)
                     {
                          if (dateToScript!= dateValues[stopPosition])
                          break;
                     }
                     break; 
                  }
                stopPosition++;
            }
          stopPosition++;
         }
        else { startPosition =1; stopPosition = sheetToScript.getLastRow();
         Browser.msgBox("Whole month");}
  
         var countRows = stopPosition - startPosition -1 ;
         if (countRows <=0) { Browser.msgBox("No list with selected date"); return 0 ; }
      
            

        var wholeTable = sheetToScript.getRange(startPosition+1, 1,countRows , lastColumn);
        var data = wholeTable.getValues();
        var rowsToScript =  getRowsToScript(data,isWholeMonth, dateToScript)
        var dateColumn = 0, commentColumn= 5, amountOfLeadsColumn = 2, linkColumn = 4, statusColumn = 7, dateSciptColumn = 8, scriptColumn = 9, readColumn = 10;

        for (row in rowsToScript)
        {
         
              var currentRow = rowsToScript[row];
              data[currentRow][readColumn] = "read";
              //wholeTable.setValues(data);
             var result = cleanTheList(data[currentRow][linkColumn]);
            if (result == true) {
                data[currentRow][scriptColumn] = "done"; 
                data[currentRow][dateSciptColumn] = currentDate; 
                
                 if (checkRejectionRate) getRejectionRate(data[currentRow][linkColumn].toString(),data[currentRow][dateColumn].toString());
                
                }
            else if (result == false) data[currentRow][commentColumn] =  "unChecked" ;
            else  data[currentRow][scriptColumn] =  result ;
            
             wholeTable.setValues(data)
        

       
    
    }
        //wholeTable.setValues(data);

        
   
    
    }
    
    
    else Browser.msgBox(nameOfSheet + " is missing");
  //Browser.msgBox("Script is Done");
}
function getRowsToScript(data, isWholeMonth, dateToScript)
{
     var masRows = []
     var dateColumn = 0, commentColumn= 5, amountOfLeadsColumn = 2, linkColumn = 4, statusColumn = 7, dateSciptColumn = 8, scriptColumn = 9, readColumn = 10;
     if (!isWholeMonth)
     {
       for (var i = 0; i < data.length; i++) {
          
         // Browser.msgBox(dateToScript + " "+ data[i][dateColumn]);
         
            if ( data[i][dateColumn] != dateToScript) continue ;
            if (data[i][scriptColumn] == "done" || data[i][linkColumn] == 0) continue;
            var status = data[i][statusColumn].toString().toLowerCase();
            var comment = data[i][commentColumn].toString().toLowerCase();

            if ( data[i][amountOfLeadsColumn] > 50  )
               if ( (status != "done" && comment == "") || (comment != "platform" && status == "" ))  continue; 
            if (data[i][commentColumn].toString().toLowerCase()== "no db") continue;
            
            masRows.push(i);
    
        }
     }
     else {
          //var count = 0;
          
          for (var i = 0; i < data.length; i++) {
              //if (count>100) break;
              if (data[i][scriptColumn] == "done" || data[i][linkColumn] == 0) continue;
              if (data[i][commentColumn].toString().toLowerCase()== "no db") continue;
              masRows.push(i);  
             // count++;
              
          }
     }
     Browser.msgBox("Rows to script: " + masRows.length)
     return masRows
}
function getRejectionRate(link,date)
{
try {
    var rejSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1rYysrmVoO4bT9h7lgeyZXr8oAGbc9jLqsyxnzCINSg8/edit#gid=0").getSheets()[0];
    var sh = SpreadsheetApp.openByUrl(link);
   
    var curSheet = SpreadsheetApp.openByUrl(link).getSheets()[0];
    var columns = curSheet.getRange(1,1,1,curSheet.getLastColumn()).getValues();
    

    var titleColumn = columns[0].indexOf("title")+1;
    var countryColumn = columns[0].indexOf("country")+1;
    var employeesColumn = columns[0].indexOf("employees")+1;
    var revenueColumn = columns[0].indexOf("revenue")+1;
    var industryColumn = columns[0].indexOf("industry")+1;
    var companyColumn = columns[0].indexOf("company")+1;
    var lastRow = curSheet.getLastRow();
    
    var countChecked = 0, countRejTitlesGreen = 0, countRejTitlesYellow = 0, countRejCountryGreen = 0,countRejCountryYellow = 0, countRejIndustryGreen = 0,countRejIndustryYellow = 0, 
          countRejEmployeesGreen = 0,countRejEmployeesYellow = 0, countRejRevenueGreen = 0,countRejRevenueYellow = 0, countRejNAC = 0;
          
    //var OV_comments = curSheet.getRange(2,OV_CommentColumn,curSheet.getLastRow()).getValues();
    var countChecked = getCheckedLeads(curSheet);
    if (countChecked==-1) {Browser.msgBox("ov_comment missing"); return ; }
    var countsGreenAndOthers = getRejCounts(curSheet, titleColumn);
    var RejTitlesGreen = countsGreenAndOthers[0];
    countRejTitlesGreen = (countsGreenAndOthers[0]/countChecked*100).toFixed(2), countRejTitlesYellow = (countsGreenAndOthers[1]/countChecked*100).toFixed(2);
    var countsGreenAndOthers = getRejCounts(curSheet, countryColumn);
    var rejCountryGreen = countsGreenAndOthers[0];
    countRejCountryGreen = (countsGreenAndOthers[0]/countChecked*100).toFixed(2), countRejCountryYellow = (countsGreenAndOthers[1]/countChecked*100).toFixed(2);
    var countsGreenAndOthers = getRejCounts(curSheet, industryColumn);
    countRejIndustryGreen = (countsGreenAndOthers[0]/countChecked*100).toFixed(2), countRejIndustryYellow = (countsGreenAndOthers[1]/countChecked*100).toFixed(2);
    var rejIndustryGreen = countsGreenAndOthers[0];
    var countsGreenAndOthers = getRejCounts(curSheet, employeesColumn);
    countRejEmployeesGreen = (countsGreenAndOthers[0]/countChecked*100).toFixed(2), countRejEmployeesYellow = (countsGreenAndOthers[1]/countChecked*100).toFixed(2);
    var rejEmployeesGreen = countsGreenAndOthers[0] ;
    var countsGreenAndOthers = getRejCounts(curSheet, revenueColumn);
    var rejRevenueGreen = countsGreenAndOthers[0];
    countRejRevenueGreen = (countsGreenAndOthers[0]/countChecked*100).toFixed(2), countRejRevenueYellow = (countsGreenAndOthers[1]/countChecked*100).toFixed(2);
    var rejNac = getRejNac(curSheet, companyColumn);
    var countRejNAC = (rejNac/countChecked*100).toFixed(2);
    //countRejNAC = ((countsGreenAndOthers[0] + countsGreenAndOthers [1])/countChecked*100).toFixed(2);
   // var date =  new Date(); 
    var name = sh.getName();
    var lastCellDate = rejSheet.getRange(rejSheet.getLastRow(),1);
    //ead1dc - light magenta 3
    //d9d2e9 - light purple 3
       
    //#cfe2f3 - default (light blue)

    var newBgColor = "#cfe2f3"
    if (lastCellDate.getValue() !=date) {
        if ( lastCellDate.getBackground() == "#ead1dc")
            newBgColor = "#d9d2e9";
        else newBgColor = "#ead1dc";
    }
    else newBgColor = lastCellDate.getBackground();
    
    rejSheet.appendRow([date,name,link,countRejTitlesGreen+'%',RejTitlesGreen, countRejTitlesYellow+'%',countRejCountryGreen+'%',rejCountryGreen, countRejCountryYellow+'%',countRejIndustryGreen+'%',rejIndustryGreen, countRejIndustryYellow+'%',countRejEmployeesGreen+'%',rejEmployeesGreen,countRejEmployeesYellow+'%',countRejRevenueGreen+'%',rejRevenueGreen, countRejRevenueYellow+'%',countRejNAC+'%',rejNac,countChecked])
    rejSheet.getRange(rejSheet.getLastRow(),1,1,3).setBackground(newBgColor)
    
    }
    catch (err) {Browser.msgBox("in rejection function " + err)}

}
function getCheckedLeads (curSheet )
{
    var OV_comments = ["Y1: linkedin/company website","Y2: PL Summary","Y3: Facebook","Y4: Suspicious Linkedin","Y5: 3rd Party Prooflink","N1: NWC","N2: Out of Business/Bad data","N/A: PV Tool","N/A: Title/PL Summary","N/A: Industry","N/A: Emp. Size","N/A: Revenue","N/A: Probably NWC",
    "N/A: Country/GEO","N/A: NAC/SUP","N/A: NAC","NAC","N/A: Prooflink","N/A: Back-up (verified)","N/A: Wrong email/General domain","N/A: Other","Q1: Questionable Title","Q2: Questionable Company","Q3: Other","N/A: Country"];
    //Browser.msgBox(OV_comments)
    var columns = curSheet.getRange(1,1,1,curSheet.getLastColumn()).getValues();
    var OV_CommentColumn = columns[0].indexOf("ov_comment")+1;
    
    var OV_comments_all = curSheet.getRange(2,OV_CommentColumn,curSheet.getLastRow()).getValues();
    if (OV_CommentColumn <=0) return -1;
    var countCheckedLeads = 0;
    OV_comments_all.forEach(function(cell) {
   // Browser.msgBox(cell)
        if (OV_comments.indexOf(cell.toString())!=-1)
            countCheckedLeads++;
    });
    return countCheckedLeads;
}
function getRejNac(curSheet, columnIndex)
{

    var countRejection = 0;
    if (columnIndex <=0) return 0;
    var columns = curSheet.getRange(1,1,1,curSheet.getLastColumn()).getValues();
    var OV_CommentColumn = columns[0].indexOf("ov_comment")+1;
    
    var OV_Comments = curSheet.getRange(2,OV_CommentColumn,curSheet.getLastRow()).getValues();
    for (var i=0; i<OV_Comments.length; i++)
    {
        if (OV_Comments[i].toString().indexOf('NAC')!=-1 ||OV_Comments[i].toString().indexOf('SUP')!=-1)
        countRejection++;
    }
    //Browser.msgBox(countRejection)
    return countRejection;
}
function getRejCounts(curSheet, columnIndex)
{
    var countRejectionGreen=0, countRejectionYellow=0;
//Browser.msgBox(columnIndex);
    if (columnIndex <=0) return 0;
    var rangeColumn =  curSheet.getRange(2,columnIndex,curSheet.getLastRow());
    var fontColors = rangeColumn.getFontColors();
    var bgColors = rangeColumn.getBackgrounds();
    
    for (var i=0; i<fontColors.length; i ++)
    {
        if (fontColors[i] == "#ff0000")
        {
            if (bgColors[i] == "#93c47d" )
              countRejectionGreen++;
            else countRejectionYellow++;
        }
       
    }
    
    return [countRejectionGreen,countRejectionYellow];
}
function getNameOfSheet(MonthSelected, YearSelected) { //not work yet
    var nameOfSheet = "";
    var months = ["Jan","Feb","March","April","May","June","July","Aug","Sep","Oct","Nov","Dec" ];
    var year ="" +YearSelected ;
    year = year.replace("20","");
    return ( months[MonthSelected-1] + "_" + year );  
}
function replaceSpaces(text)
{
     text  = text.toString().toLowerCase();
     text = text.replace(/^\s*/, '').replace(/\s*$/, '');
     return text;
}
function generateCurrentDate() {

    var today = new Date();
    var dd = today.getDate();
    var mm = today.getMonth() + 1; 
    var yyyy = today.getFullYear();

    if (dd < 10) {
        dd = '0' + dd
    }

    if (mm < 10) {
        mm = '0' + mm
    }

    today = mm + '/' + dd + '/' + yyyy;
    return today
}

