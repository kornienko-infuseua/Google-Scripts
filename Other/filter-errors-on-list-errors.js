function onOpen() {
 
 PropertiesService.getScriptProperties().setProperty('StartRow', 0);

  var subSubMenu = SpreadsheetApp.getUi().createMenu('Join-Split')
    .addItem('Join', 'join')
    .addItem('Split', 'split')
    .addItem('Delete Quotes', 'deleteQuotes')
    .addItem('Add Quotes', 'addQuotes');
  SpreadsheetApp.getUi()
    .createMenu('Script')
    .addItem('Filter errors from CSV', 'fromHtml').addSeparator()
    //.addItem('test','someTest')
   // .addItem('someTest', 'someTest').addSeparator()
   // .addItem('delete triggers', 'deletetriggers').addSeparator()
 //   .addItem('Test3', 'sometest3').addSeparator()
    //.addItem("Get Changed Industries", "getIndustries").addSeparator()
    .addSubMenu(subSubMenu)
    .addToUi();
}
function someTest() {
try{
    var sheet = SpreadsheetApp.getActiveSheet();
    //var industries = ["Advertisement / Marketing","Aerospace / Aviation","Agriculture","Automotive","Biotech and Pharmaceuticals","Computers and Technology","Construction","Corporate Services","Education","Finance","Government","Healthcare / Medical","Industry","Insurance","Legal","Manufacturing","Media","Non-Profit / Organizations","Real Estate","Retail and Consumer Goods","Service Industry","Telecommunication","Transportation and Logistics","Travel / Hospitality / Entertainment","Utility / Energy"]
   var subindustries = ["Accounting","Advertising Agencies","Aerospace / Defense - Major Diversified","Aerospace / Defense Products & Services","Agriculture","Air Delivery & Freight Services","Air Services","Alternative Dispute Resolution","Alternative Dispute Resolution","Alternative Medicine","Animation","Apparel / Fashion","Architecture / Planning","Arts / Crafts","Auto Dealership","Auto Manufacturers","Auto Parts","Auto Parts Stores","Auto Parts Wholesale","Banking","Biotech","Broadcast Media","Broadcasting - Radio","Broadcasting - TV","Building Materials","Business Supplies / Equipment","Business Supplies / Equipment","Cable and Other Program Distribution","Capital Markets","CATV Systems","Chemicals","Civic / Social Organization","Civil Engineering","Clinical Laboratory","Commercial Real Estate","Communication Equipment","Communications","Computer / Network Security","Computer Games","Computer Hardware","Computer Networking","Computer Software","Construction","Consumer Electronics","Consumer Electronics","Consumer Goods","Consumer Goods","Consumer Services","Cosmetics","Dairy","Defense / Space","Design","Diversified Communication Services","eCommerce","Education Management","E-Learning","E-Learning","Electrical / Electronic Manufacturing","Electrical Power","Entertainment","Entertainment","Environmental Services","Event Management","Events Services","Executive Office","Facilities Services","Farming","Financial Services","Fine Art","Fishery","Food / Beverages","Food / Beverages","Food Production","Food Production","Fundraising","Funeral Service and Crematories","Furniture","Furniture","Gambling / Casinos","Gas","Glass / Ceramics / Concrete","Glass / Ceramics / Concrete","Government Administration","Government Relations","Graphic Design","Health","Health / Wellness / Fitness","Higher Education","Hospital / Health Care","Hospitality","Human Resources","Import / Export","Individual / Family Services","Industrial Automation","Industrial Automation","Information Services","Information Technology / Services","Internation Trade / Development","International Affairs","International Trade / Development","Internet","Investment Banking / Venture","Investment Management","Judiciary","Judiciary","Law Enforcement","Law Enforcement","Law Practice","Legal Services","Legislative Office","Legislative Office","Leisure / Travel","Libraries","Libraries","Life","Life Sciences Manufacturing","Logistics / Supply Chain","Long Distance Carriers","Luxury Goods / Jewelry","Machinery","Major Airlines","Management Consulting","Manufacturing Retail","Maritime","Market Research","Marketing Services","Mechanical / Industrial Engineering","Media Production","Medical Device","Medical Practice","Mental Health Care","Military","Mining / Metals","Mortgage","Motion Pictures / Film","Movie Production / Theaters","Museums / Institutions","Museums / Institutions","Music","Nanotechnology","Network Functions NAC","Newspapers","Non-Profit Organization Management","Nuclear","Oil","Online Publishing","Outsourcing / Offshoring","Package / Freight Delivery","Package / Freight Delivery","Packaging / Containers","Packaging / Containers","Paging","Paper / Forest Products","Performing Arts","Performing Arts","Petroleum","Pharmaceuticals","Philanthropy","Photography","Plastics","Political Organization","Political Organization","Power Distributors","Power Generation","PR","Primary / Secondary","Printing","Printing","Processing Systems / Products","Professional Training","Program Development","Program Development","Property Management","Public Policy","Public Safety","Publishing","Publishing","Ranching","Recreational Facilities / Services","Recreational Vehicles","Regional Airlines","Religious Institutions","Renewable Energy","Research","Research","Residential Real Estate","Restaurants","Restaurants","Retail","Satellite Telecommunications","Security / Investigations","Security / Investigations","Semiconductors","Shipbuilding","Sporting Goods","Sports","Staffing / Recruiting","Sub-Industry","Supermarkets","Telecom Services - Domestic","Telecom Services - Foreign","Textiles","Think Tanks","Think Tanks","Tires","Tobacco","Translation / Localization","Translation / Localization","Transportation / Trucking / Railroad","Trucks / Other Vehicles","VAR/VAD/System Integrators","Vehicle","Venture Capital","Veterinary","Warehousing","Waste Management","Water Treatment","Wholesale","Wine / Spirits","Wine / Spirits","Wired Telecommunications Carriers","Wireless Communications","Writing / Editing"]
   var range = sheet.getRange(2,3,sheet.getLastRow()-1)
    var values = range.getValues();
    var green = "#cfe2f3";
    for (var i=0 ;i< values.length; i++)
    {
        if (subindustries.indexOf(values[i][0].toString())!=-1)
            sheet.getRange(i+2,3).setBackground(green)
    }
    }
    catch (err)
    {
        Browser.msgBox(err)
    }

}
function doWork() {

}
function sometest3() {
  var triggerDay = (new Date()).getTime();
  triggerDay = new Date(triggerDay + 600000)
  ScriptApp.newTrigger("myFunction")
    .timeBased()
    .at(triggerDay)
    .create();

}
function myFunction() {
  Browser.msgBox('my function 3')

}
function deletetriggers() {

  //Browser.msgBox(triggers.length)
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }


}

function runMe() {


  //Browser.msgBox('runMe')
  var MAX_RUNNING_TIME = 10000;
  var startTime = (new Date()).getTime();
  var REASONABLE_TIME_TO_WAIT = 10;

  var scriptProperties = PropertiesService.getScriptProperties();
  var startRow = scriptProperties.getProperty('StartRow');
  startRow = 0;
  for (var ii = startRow; ii <= 1000000000; ii++) {
    var currTime = (new Date()).getTime();
    if (currTime - startTime >= MAX_RUNNING_TIME) {
      //Browser.msgBox('exceeded time')
      scriptProperties.setProperty("StartRow", ii);

      ScriptApp.newTrigger("function1")
        .timeBased()
        .everyMinutes(1)
        .create();

      // var triggers = ScriptApp.getProjectTriggers();
      // Browser.msgBox(triggers[0].getHandlerFunction())
      break;
    } else {
      doSomeWork();
    }
  }

  //do some more work here

}
function doSomeWork() {


}
function function1() {

  Browser.msgBox("function 1")
}
//function onEdit(e) {
//var ss = SpreadsheetApp.getActiveSpreadsheet();
//var currentSheet = ss.getActiveSheet();
// var nameOfSheet = currentSheet.getName();
//var upd = e.range.getValue();
//if (upd && nameOfSheet == "Leads errors") 
//}
function fromHtml() {

  var dialog = HtmlService.createHtmlOutputFromFile('form.html').setHeight(150).setWidth(300);
  SpreadsheetApp.getUi().showModalDialog(dialog, "Select File");
}

function serverFunc(theForm) {

  var anExampleText = theForm.anExample;  // This is a string
  var fileBlob = theForm.theFile;

  var rows = fileBlob.contents.split('\n')// This is a Blob.

  var startRow = PropertiesService.getScriptProperties().getProperty('StartRow');
  if (!(startRow > 0)) startRow = 0;
  processRowsFromFile(rows, startRow)
}

function processRowsFromFile(rows, startRow) {
  var Properties =   PropertiesService.getScriptProperties();
  var sheetToAppendCompany = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Company errors');
  var sheetToAppendLead = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads errors');
  var sheetToAppendCountry = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('New countries');

  var link = "";
  var curSheet;
  var currentDate = generateCurrentDate();
  var LeadsArray = [], CompaniesArray = [], CountriesArray = []
  var startTime = (new Date()).getTime();
  
  var r = 1 + parseInt(startRow);
  var max_r = rows.length - 1;
  
  for  ( ; r<max_r; r++ ) {
    try {
    
      var currTime = (new Date()).getTime();
      if ((currTime - startTime) > 300000) {
        
        Properties.setProperty('StartRow', r);
        
        break;
      }
     

      else {

        var errorInfo = getErrorInfo(rows[r])

        if (!errorInfo) continue;

        if (errorInfo.link != link) { //if prooflink is the same as in preious row, don't open spreadsheet again
          link = errorInfo.link
          try {
            curSheet = SpreadsheetApp.openByUrl(link).getSheets()[0];
          } catch (err) { continue }
        }

        if (errorInfo.errorType == 3) {
          var leadInfo = getLeadInfo(curSheet, errorInfo.row);

          if (leadInfo != false)
            LeadsArray.push([currentDate, errorInfo.date, errorInfo.errorMessage, leadInfo.first_name, leadInfo.last_name, leadInfo.companyName, leadInfo.email, leadInfo.prooflink, errorInfo.row, errorInfo.link])
        }
        else if (errorInfo.errorType == 4) {
          CountriesArray.push([currentDate, errorInfo.date, errorInfo.errorMessage, errorInfo.countryName, errorInfo.row, errorInfo.link])

        }
        else {
          var compInfo = getCompanyInfo(curSheet, errorInfo.row, errorInfo.errorType)
          CompaniesArray.push([errorInfo.date, errorInfo.errorMessage, errorInfo.oldName ? errorInfo.oldName : compInfo.companyName, errorInfo.newName ? errorInfo.newName : "", compInfo.domain, errorInfo.row, errorInfo.link, compInfo.EmployeesProoflink ? compInfo.EmployeesProoflink : "", compInfo.Employees ? compInfo.Employees : "", compInfo.RevenueProoflink ? compInfo.RevenueProoflink : "", compInfo.Revenue ? compInfo.Revenue : "", compInfo.prooflink])

        }

      }
     
    } catch (err) { //Browser.msgBox('err in row : ' + r + "   " + err); 
    continue; }
  }

  
  try {

    if (CompaniesArray.length > 0) sheetToAppendCompany.getRange(sheetToAppendCompany.getLastRow() + 1, 1, CompaniesArray.length, CompaniesArray[0].length).setValues(CompaniesArray);
    if (LeadsArray.length > 0) sheetToAppendLead.getRange(sheetToAppendLead.getLastRow() + 1, 1, LeadsArray.length, LeadsArray[0].length).setValues(LeadsArray);
    if (CountriesArray.length > 0) sheetToAppendCountry.getRange(sheetToAppendCountry.getLastRow() + 1, 1, CountriesArray.length, CountriesArray[0].length).setValues(CountriesArray);
    if ( r >= rows.length -1 ) { 
      Properties.setProperty('StartRow', 0);
      //Browser.msgBox("Script is Done")
      }
    else Browser.msgBox("Please, run script again with the same file")
  } catch (err) { Browser.msgBox('err in append   ' + err) }



}
function getLeadInfo(curSheet, row) {
  try {
    var LeadInfo = {};
    var columns = curSheet.getRange(1, 1, 1, curSheet.getLastColumn()).getValues();

    var PvColumn = columns[0].indexOf("pv_comment") + 1;
    if (PvColumn > 0) {
      var pvComment = curSheet.getRange(row, PvColumn).getValue();
      if (pvComment.indexOf("NWC") != -1) { return false; }
    }
    var CompanyColumn = columns[0].indexOf("company") + 1;
    if (CompanyColumn == 0) return false;
    LeadInfo.companyName = curSheet.getRange(row, CompanyColumn).getValue();
    var prooflink = columns[0].indexOf("prooflink") + 1;
    if (prooflink)
      LeadInfo.prooflink = curSheet.getRange(row, prooflink).getValue();
    else LeadInfo.prooflink = ""
    var first_name = columns[0].indexOf("first_name") + 1;
    if (first_name)
      LeadInfo.first_name = curSheet.getRange(row, first_name).getValue();
    else LeadInfo.first_name = ""
    var last_name = columns[0].indexOf("last_name") + 1;
    if (last_name)
      LeadInfo.last_name = curSheet.getRange(row, last_name).getValue();
    else LeadInfo.last_name = ""
    var email = columns[0].indexOf("email") + 1;
    if (email)
      LeadInfo.email = curSheet.getRange(row, email).getValue();
    else LeadInfo.email = ""
    return LeadInfo;
  } catch (err) { Browser.msgBox('err getLeadInfo    ' + err) }
}
function getErrorInfo(str) {
  try {

    var info = {}
    if (!str) { return; }
    var columnsArr = str.split(',')
    info.link = columnsArr[4];
    info.row = columnsArr[3];
    if (str.search("Company is changed") >= 0) {
      var dbName = columnsArr[2].substring(30);
      var pos = dbName.substring(0, dbName.search('"'))
      if (pos != -1) {
        info.errorType = 0; //company is changed
        info.oldName = pos;
        var listName = columnsArr[3].substring(12);
        info.newName = listName.substring(0, listName.search('"'))

        info.link = columnsArr[5];
        info.row = columnsArr[4];
        info.errorMessage = "Company is changed"
      }
    }

    else if (columnsArr[2].search('Employees PL') >= 0) {
      info.errorMessage = "New employees for VM company"
      info.errorType = 1;  // 1 - employees changed
    }
    else if (columnsArr[2].search('Revenue PL') >= 0) {
      info.errorMessage = "New revenue for VM company"
      info.errorType = 2;  // 2 - revenue changed
    }
    else if (columnsArr[2].search('bad data') >= 0) {
      info.errorMessage = "Contact is marked as bad data"
      info.errorType = 3;  // 3 - contact is marked as bad data
    }
    else if (columnsArr[2].search('Country not') >= 0) {
      info.errorMessage = "Country not found";
      info.errorType = 4; // 4 - countrry not found
      var countryPos = columnsArr[2].search('Country not found');
      //Browser.msgBox(columnsArr[2])
      var lastSymbol = columnsArr[2][columnsArr[2].length-1] 
      info.countryName = columnsArr[2].substring(countryPos + 19, lastSymbol=='"'? columnsArr[2].length-1 : columnsArr[2].length)
    }
    else { return 0; }
    info.date = columnsArr[0];

    return info;
  } catch (err) { Browser.msgBox('err getErrorInfo  ' + err) }
}
function getCompanyInfo(curSheet, row, changedType) {
  try {
    var compInfo = {}

    var columns = curSheet.getRange(1, 1, 1, curSheet.getLastColumn()).getValues();


    if (changedType == 1 || changedType == 2) {
      var CompanyColumn = columns[0].indexOf("company") + 1;
      if (CompanyColumn == 0) return false;
      compInfo.companyName = curSheet.getRange(row, CompanyColumn).getValue();
      var columnProoflinkName, columnValueName;
      if (changedType == 1) {
        columnProoflinkName = "employees_prooflink";
        columnValueName = "employees";
      } else {
        columnProoflinkName = "revenue_prooflink";
        columnValueName = "revenue";
      }
      var columnValue = columns[0].indexOf(columnValueName) + 1;
      var columnProoflink = columns[0].indexOf(columnProoflinkName) + 1;


      if (columnProoflink && curSheet.getRange(row, columnProoflink).getBackground() == '#ffff00') {
        if (changedType == 1)
          compInfo.EmployeesProoflink = curSheet.getRange(row, columnProoflink).getValue();
        else compInfo.RevenueProoflink = curSheet.getRange(row, columnProoflink).getValue();
      }
      if (columnValue && curSheet.getRange(row, columnValue).getBackground() == '#ffff00') {
        if (changedType == 1)
          compInfo.Employees = curSheet.getRange(row, columnValue).getValue();
        else compInfo.Revenue = curSheet.getRange(row, columnValue).getValue();
      }
    }
    else { // нужно переделать, г**код
      var columnValueEmployees = columns[0].indexOf("employees") + 1;
      var columnEmployeeesProoflink = columns[0].indexOf("employees_prooflink") + 1;
      var columnValueRevenue = columns[0].indexOf("revenue") + 1;
      var columnRevenueProoflink = columns[0].indexOf("revenue_prooflink") + 1;

      if (columnValueEmployees && curSheet.getRange(row, columnValueEmployees).getBackground() == '#ffff00') {
        compInfo.Employees = curSheet.getRange(row, columnValueEmployees).getValue();
      }
      if (columnEmployeeesProoflink && curSheet.getRange(row, columnEmployeeesProoflink).getBackground() == '#ffff00') {
        compInfo.EmployeesProoflink = curSheet.getRange(row, columnEmployeeesProoflink).getValue();
      }
      if (columnValueRevenue && curSheet.getRange(row, columnValueRevenue).getBackground() == '#ffff00') {
        compInfo.Revenue = curSheet.getRange(row, columnValueRevenue).getValue();
      }
      if (columnRevenueProoflink && curSheet.getRange(row, columnRevenueProoflink).getBackground() == '#ffff00') {
        compInfo.RevenueProoflink = curSheet.getRange(row, columnRevenueProoflink).getValue();
      }
    }
    var emailColumn = columns[0].indexOf("email") + 1;
    if (emailColumn == 0) return false;
    var email = curSheet.getRange(row, emailColumn).getValue();
    compInfo.domain = email.substring(email.search('@') + 1);

    var prooflink = columns[0].indexOf("prooflink") + 1;
    if (prooflink)
      compInfo.prooflink = curSheet.getRange(row, prooflink).getValue();
    else compInfo.prooflink = ""
    return compInfo;


  } catch (err) {// Browser.msgBox('err getCompanyInfo     ' + err) 
  }
}
function getIndustries() {
  try {
    var rejectionSheet = SpreadsheetApp.getActiveSpreadsheet();
    var sh = rejectionSheet.getSheetByName('New lists');
    var sheetIndustryChanged = rejectionSheet.getSheetByName('IndustryChanged')
    var linkValues = sh.getRange(2, 4, sh.getLastRow() - 1).getValues()
    //Browser.msgBox(linkValues)
    var count = 0;
    for (var i = 0; i < linkValues.length; i++ , count++) {
      if (count == 5) { Browser.msgBox("Click me"); count++; }
      var link = linkValues[i].toString()
      var arrInfo = getCompaniesIndustries(link);
      var rowstoAppend = []
      if (arrInfo) {
        for (i in arrInfo) {
          rowstoAppend.push([arrInfo[i].companyName, arrInfo[i].newIndustry, arrInfo[i].row, link])
        }
        // Browser.msgBox(rowstoAppend)
      }
    }
    for (i in rowstoAppend) sheetIndustryChanged.appendRow(rowstoAppend[i])
    //Browser.msgBox("ROWS: " +rowstoAppend)
  } catch (err) { Browser.msgBox('err get industries  ' + err) }
}

function getCompaniesIndustries(link) {
  var industries = ["Advertisement / Marketing", "Aerospace / Aviation", "Agriculture", "Automotive", "Biotech and Pharmaceuticals", "Computers and Technology", "Construction", "Corporate Services", "Education", "Finance", "Government", "Healthcare / Medical", "Industry", "Insurance", "Legal", "Manufacturing", "Media", "Non-Profit / Organizations", "Real Estate", "Retail and Consumer Goods", "Service Industry", "Telecommunication", "Transportation and Logistics", "Travel / Hospitality / Entertainment", "Utility / Energy"]
  var curSheet;
  try {
    curSheet = SpreadsheetApp.openByUrl(link).getSheets()[0];
  } catch (err) { return false }
  var columns = curSheet.getRange(1, 1, 1, curSheet.getLastColumn()).getValues();
  var industryColumn = columns[0].indexOf("industry") + 1;
  var companyNameColumn = columns[0].indexOf("company") + 1;
  var companyVerifiedColumn = columns[0].indexOf("company_verified") + 1;
  if (!industryColumn || !companyNameColumn || !companyVerifiedColumn) return false;

  var arrInfo = [];
  for (var i = 2; i < curSheet.getLastRow() - 1; i++) {
    if (curSheet.getRange(i, industryColumn).getBackground() == "#ffff00" && curSheet.getRange(i, companyNameColumn).getBackground() != "#ffff00" && industries.indexOf(curSheet.getRange(i, industryColumn).getValue()) != -1 && curSheet.getRange(i, companyVerifiedColumn).getValue() == "VM")
      arrInfo.push({ companyName: curSheet.getRange(i, companyNameColumn).getValue(), newIndustry: curSheet.getRange(i, industryColumn).getValue(), row: i })

  }


  return arrInfo;

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