function onOpen() {
    var subSubMenu = SpreadsheetApp.getUi().createMenu('Join-Split')
        .addItem('Join', 'join')
        .addItem('Split', 'split')
        .addItem('Delete Quotes', 'deleteQuotes')
        .addItem('Add Quotes', 'addQuotes');
    SpreadsheetApp.getUi()
        .createMenu('Script')
        .addItem('Filter errors from CSV', 'fromHtml').addSeparator()
        //.addItem("Get Changed Industries", "getIndustries").addSeparator()
        .addSubMenu(subSubMenu)
        .addToUi();
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

    processRowsFromFile(rows)
    //return adoc.getUrl();
}

function processRowsFromFile(rows) {
    var sheetToAppendCompany = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Company errors');
    var sheetToAppendLead = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads errors');
    var sheetToAppendCountry = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('New countries');

    var link = "";
    var curSheet;
    var currentDate = generateCurrentDate();
    var LeadsArray = [], CompaniesArray = [], CountriesArray = []

    for (var r = 1, max_r = rows.length - 1; r < max_r; ++r) {
        try {
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

        } catch (err) { Browser.msgBox(err); continue; }
    }


    try {
        if (CompaniesArray.length > 0) sheetToAppendCompany.getRange(sheetToAppendCompany.getLastRow() + 1, 1, CompaniesArray.length, CompaniesArray[0].length).setValues(CompaniesArray);
        if (LeadsArray.length > 0) sheetToAppendLead.getRange(sheetToAppendLead.getLastRow() + 1, 1, LeadsArray.length, LeadsArray[0].length).setValues(LeadsArray);
        if (CountriesArray.length > 0) sheetToAppendCountry.getRange(sheetToAppendCountry.getLastRow() + 1, 1, CountriesArray.length, CountriesArray[0].length).setValues(CountriesArray);
    } catch (err) { Browser.msgBox(err) }


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
    } catch (err) { Browser.msgBox(err) }
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
            //var countryPos = columnsArr[2].search('Country not found');
            info.countryName = columnsArr[2].substring(20, columnsArr[2].length )
        }
        else { return 0; }
        info.date = columnsArr[0];

        return info;
    } catch (err) { Browser.msgBox(err) }
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


    } catch (err) { Browser.msgBox(err) }
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
    } catch (err) { Browser.msgBox(err) }
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
function getDomain(sheet, row) {
    try {
        var columns = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
        var emailColumn = columns[0].indexOf("email") + 1;
        if (emailColumn == 0) return false;
        var email = sheet.getRange(row, emailColumn).getValue();

        return email;
        //return email.substring(email.search('@') + 1);

    } catch (err) { Browser.msgBox(err) }
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