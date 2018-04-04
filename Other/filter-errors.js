function onOpen() {
    var subSubMenu = SpreadsheetApp.getUi().createMenu('Join-Split')
        .addItem('Join', 'join')
        .addItem('Split', 'split')
        .addItem('Delete Quotes', 'deleteQuotes')
        .addItem('Add Quotes', 'addQuotes');
    SpreadsheetApp.getUi()
        .createMenu('Script')
        .addItem('Get Changed companies', 'fromHtml').addSeparator()
        .addItem("Get Changed Industries", "getIndustries").addSeparator()
        .addSubMenu(subSubMenu)
        .addToUi();
}


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
    var sheetToAppend = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CompanyChanged')
    var link = "";
    var curSheet;
    var currentDate = generateCurrentDate();
    for (var r = 1, max_r = rows.length; r < max_r; ++r) {
        try {
            var str = rows[r];
            if (str.search('Company is changed') == -1) continue;

            var errorInfo = getErrorInfo(str)
            if (errorInfo.link != link) {
                link = errorInfo.link


                try {
                    curSheet = SpreadsheetApp.openByUrl(link).getSheets()[0];

                } catch (err) { continue }
            }

            errorInfo.domain = getDomain(curSheet, errorInfo.row)
            //Browser.msgBox(errorInfo.date + ' ' + errorInfo.link + ' ' + errorInfo.row + ' ' + errorInfo.oldName + ' ' + errorInfo.newName + ' ' + errorInfo.domain)
            sheetToAppend.appendRow([currentDate, errorInfo.date, errorInfo.oldName, errorInfo.newName, errorInfo.domain, errorInfo.row, errorInfo.link])
        } catch (err) { continue; }
    }

    //Browser.msgBox("done")
}

function getIndustries() {
    try {
        var rejectionSheet = SpreadsheetApp.getActiveSpreadsheet();
        var sh = rejectionSheet.getSheetByName('New lists');
        var sheetIndustryChanged = rejectionSheet.getSheetByName('IndustryChanged')
        var linkValues = sh.getRange(2, 4, sh.getLastRow() - 1).getValues()
        //Browser.msgBox(linkValues)
        for (var i = 0; i < linkValues.length; i++) {
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
        Browser.msgBox("ROWS: " + rowstoAppend)
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
        return email.substring(email.search('@') + 1);

    } catch (err) { Browser.msgBox(err) }
}
function getErrorInfo(str) {
    try {
        var info = {}

        var columnsArr = str.split(',')
        info.date = columnsArr[0]
        info.link = columnsArr[5];
        info.row = columnsArr[4];
        var dbName = columnsArr[2].substring(30);
        info.oldName = dbName.substring(0, dbName.search('"'))

        var listName = columnsArr[3].substring(12);
        info.newName = listName.substring(0, listName.search('"'))
        return info;
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