function onOpen() {

    SpreadsheetApp.getUi()
        .createMenu('new Report Script')
        .addItem('DO REPORT', 'createDialogNew')
        .addSeparator()
        .addItem('upd. info', 'updateCheckers')
        .addToUi();
}

function createDialogNew() {

    var htmlDialog = HtmlService.createHtmlOutputFromFile("newScriptDialog.html")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(300)
        .setWidth(270);
    SpreadsheetApp.getUi().showModalDialog(htmlDialog, "Select Date");
}

function doReportnew(listLink, listType, dayToScript, isAppend) {
    var currentSheet = SpreadsheetApp.getActiveSheet();
    var QCInfo=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QC_Schedule");
    var userEmail = Session.getActiveUser().getEmail();
    var QC_ID;
    var QC_TabName;
    if (userEmail) {

        //var QCSchedule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QC_Schedule");
        var rangeQCID = QCInfo.getRange(4, 1, QCInfo.getLastRow(), 4).getValues();
        for (var i = 0; i < rangeQCID.length; i++)
            if (rangeQCID[i][3] == userEmail.toString()) {
                QC_ID = rangeQCID[i][2];
                QC_TabName = "QC " + rangeQCID[i][1];
                break;
            }

    }
    if (!QC_ID) { Browser.msgBox("NO QC OPERATOR"); return 0; }
    if (!listLink) {Browser.msgBox("No LINK"); return 0; }
    Browser.msgBox("Report Started. It may take some time");
    
    var result = reportList(QCInfo, listType, dayToScript, listLink, QC_ID, QC_TabName, isAppend);
}
function updateCheckers()
{
     var lvLink= "https://docs.google.com/spreadsheets/d/1R6_o3_3nDv_1_zxgSRPFztTBn-mEyBztFeT5_3KcY_w/edit#gid=62608384"
     var operTab = SpreadsheetApp.openByUrl(lvLink).getSheetByName("Oper_Tab");
     var checkersID = operTab.getRange(5, 4, operTab.getLastRow()).getValues();
     var checkersName = operTab.getRange(5, 11, operTab.getLastRow()).getValues();
     
     var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QC_Schedule")
     currentSheet.getRange(4, 6, checkersID.length).setValues(checkersID);
     currentSheet.getRange(4,7, checkersID.length).setValues(checkersName);
}
function reportList(QCInfo, listType, dayToScript, listLink, ID, QCTabName, isAppend) {
    try {
        var listToReport = SpreadsheetApp.openByUrl(listLink).getSheets()[0];
        var valuesToSearchColumn = listToReport.getRange(1, 1, 1, listToReport.getLastColumn()).getValues();

        var qcDateColumn = valuesToSearchColumn[0].indexOf("qc_date") + 1;
        var qcColumn = valuesToSearchColumn[0].indexOf("qc_comment") + 1;
        var checkedByColumn = valuesToSearchColumn[0].indexOf("checked_by") + 1;
        
        if (qcDateColumn<=0 ) {Browser.msgBox("qc_date column is missing"); return;}
        else if (qcColumn<=0 ) {Browser.msgBox("qc_comment column is missing"); return;}
        else if (checkedByColumn<=0 ) {Browser.msgBox("checked_by column is missing"); return;}
       
        

        var rowsCount = listToReport.getLastRow() - 1;

        //var rangeToSearchColumn = listToReport.getRange(1,l)
        var qcValues = listToReport.getRange(2, qcColumn, rowsCount).getValues();
        var qcValuesbgColors = listToReport.getRange(2, qcColumn, rowsCount).getFontColors();
        var qcDateValues = listToReport.getRange(2, qcDateColumn, rowsCount).getValues()
        var checkedByValues = listToReport.getRange(2, checkedByColumn, rowsCount).getValues();

        var currentDate = dayToScript;
        sortCheckers(qcDateValues, qcValues, checkedByValues, qcValuesbgColors);
        var arrCheckers = [], arrChekerStrikes = [], arrChekerStrikesString = [], arrComments = []
        countCheckers(ID, qcDateValues, qcValues, checkedByValues, qcValuesbgColors, currentDate,
            arrCheckers, arrChekerStrikes, arrChekerStrikesString, arrComments);
        if (arrCheckers.length == 0) { Browser.msgBox("Nothing to report"); return 0; }
        var dateFromRecievedList = getDateFromRecievedLists(listLink)
        if (!dateFromRecievedList) { Browser.msgBox("Missing in Recieved Lists"); return; }
        var temp = dateFromRecievedList.split('.');
        var outPutDate = new Date(temp[0] + '/' + temp[1] + '/' + temp[2]);
    }
    catch (err) { Browser.msgBox("Error from reportList " + err) }
    if (isAppend) appendRows (QCInfo, listLink, listType,
        arrCheckers, arrChekerStrikes, arrChekerStrikesString, arrComments, outPutDate, currentDate, QCTabName)
    else appendRowsUpdating(listLink, listType,
        arrCheckers, arrChekerStrikes, arrChekerStrikesString, arrComments, outPutDate, currentDate, QCTabName)
    Browser.msgBox("Report finished");
}
function appendRows (qc_info, listLink, listType,
    arrCheckers, arrChekerStrikes, arrChekerStrikesString, arrComments, outPutDate, currentDate, QCTabName)
{
try {
     var TeamQualityReportLink = "https://docs.google.com/spreadsheets/d/1R6_o3_3nDv_1_zxgSRPFztTBn-mEyBztFeT5_3KcY_w/edit#gid=1812332991";
       var TeamQualityReport = SpreadsheetApp.openByUrl(TeamQualityReportLink).getSheetByName(QCTabName);
     
      // Browser.msgBox(lastRow)
      
      var checkersID = qc_info.getRange(4, 6, qc_info.getLastRow()).getValues();
      var checkersNames = qc_info.getRange(4, 7, qc_info.getLastRow()).getValues();
        for (i in checkersID) checkersID[i] = Number(checkersID[i]);
      var dateFromRL = getDateFromRecievedLists(listLink);
      var ReportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reported");
      ReportSheet.clearContents();
      var lastRow = ReportSheet.getLastRow();
      var rangeToUpdate = ReportSheet.getRange(2,1,arrCheckers.length, 10);
      ReportSheet.getRange(1,2).setValue(QCTabName);
        //var rangeToUpdate = TeamQualityReport.getRange(lastRow, 1, arrCheckers.length, 10);
        var valuesUpdate = rangeToUpdate.getValues();
        var count = 0;
        for (var i = 0; i < arrCheckers.length; i++) {
            if (!arrCheckers[i]) continue;
            var indexChecker = checkersID.indexOf(Number(arrCheckers[i]));
            if (indexChecker == -1) continue;
            var nameChecker = checkersNames[indexChecker];
            // Browser.msgBox(nameChecker);
            valuesUpdate[count][0] = currentDate;
            valuesUpdate[count][1] = dateFromRL;
            valuesUpdate[count][2] = nameChecker;
            valuesUpdate[count][3] = SpreadsheetApp.openByUrl(listLink).getName();
            valuesUpdate[count][4] = listLink;
            valuesUpdate[count][5] = "req."
            valuesUpdate[count][6] = listType;
            valuesUpdate[count][7] = arrChekerStrikesString[i];
            valuesUpdate[count][8] = arrComments[i] + arrChekerStrikes[i];
            valuesUpdate[count][9] = arrChekerStrikes[i];
            count++;
        }
        rangeToUpdate.setValues(valuesUpdate);
        }        catch (err) { Browser.msgBox(err)}
} 
    
    
    
function appendRowsUpdating(listLink, listType,
    arrCheckers, arrChekerStrikes, arrChekerStrikesString, arrComments, outPutDate, currentDate, QCTabName) {
    try {
         Browser.msgBox("Будет очень долго............................. ")
        //Browser.msgBox("Append Rows");
        var TeamQualityReportLink = "https://docs.google.com/spreadsheets/d/1R6_o3_3nDv_1_zxgSRPFztTBn-mEyBztFeT5_3KcY_w/edit#gid=1812332991";
        var TeamQualityReport = SpreadsheetApp.openByUrl(TeamQualityReportLink).getSheetByName(QCTabName);
        var lastRow = TeamQualityReport.getLastRow()
        // Browser.msgBox(lastRow);
        var r = TeamQualityReport.getRange(6, 1, lastRow).getValues();
        var links = TeamQualityReport.getRange(6, 5, lastRow).getValues();
        var dates = TeamQualityReport.getRange(6, 1, lastRow).getValues();
        var countNotEmptyRows = 0;
        try { var lastRowIndex = r.length; } catch (err) { Browser.msgBox("lastRowIndex") }
        var curListId = SpreadsheetApp.openByUrl(listLink).getId();
        // Browser.msgBox("Before while");
        while (countNotEmptyRows < lastRowIndex) {
            if (links[countNotEmptyRows][0].indexOf(curListId) != -1) {
                //Browser.msgBox("already reported" + countNotEmptyRows)
                if (compareDates(new Date(currentDate), new Date( dates[countNotEmptyRows][0] ))) {
                    var checkerName = TeamQualityReport.getRange(6 + countNotEmptyRows, 3).getValue();
                    checker = checkerName.split(' ')[0];
                    var index = arrCheckers.indexOf(Number(checker));
                    if (index > -1) {
                        var updateCheckerRange = TeamQualityReport.getRange(6 + countNotEmptyRows, 8, 2, 3);
                        var updateValues = updateCheckerRange.getValues();
                        updateValues[0][0] = arrChekerStrikesString[index];
                        updateValues[0][1] = arrComments[index] + arrChekerStrikes[index];
                        updateValues[0][2] = arrChekerStrikes[index];
                        updateCheckerRange.setValues(updateValues);
                        arrCheckers[index] = null;
                    }
                }
            }
            if (r[countNotEmptyRows] != "") countNotEmptyRows++;
            else break;
        }
        //Browser.msgBox("after while");
        var dateFromRL = getDateFromRecievedLists(listLink);
        var operTab = SpreadsheetApp.openByUrl(TeamQualityReportLink).getSheetByName("Oper_Tab");
        var checkersID = operTab.getRange(5, 4, operTab.getLastRow()).getValues();
        var checkersNames = operTab.getRange(5, 11, operTab.getLastRow()).getValues();
        for (i in checkersID) checkersID[i] = Number(checkersID[i]);
        var rangeToUpdate = TeamQualityReport.getRange(countNotEmptyRows + 6, 1, arrCheckers.length, 10);
        var valuesUpdate = rangeToUpdate.getValues();
        var count = 0;
        for (var i = 0; i < arrCheckers.length; i++) {
            if (!arrCheckers[i]) continue;
            var indexChecker = checkersID.indexOf(Number(arrCheckers[i]));
            if (indexChecker == -1) continue;
            var nameChecker = checkersNames[indexChecker];
            // Browser.msgBox(nameChecker);
            valuesUpdate[count][0] = currentDate;
            valuesUpdate[count][1] = dateFromRL;
            valuesUpdate[count][2] = nameChecker;
            valuesUpdate[count][3] = SpreadsheetApp.openByUrl(listLink).getName();
            valuesUpdate[count][4] = listLink;
            valuesUpdate[count][5] = "req."
            valuesUpdate[count][6] = listType;
            valuesUpdate[count][7] = arrChekerStrikesString[i];
            valuesUpdate[count][8] = arrComments[i] + arrChekerStrikes[i];
            valuesUpdate[count][9] = arrChekerStrikes[i];
            count++;
        }
        rangeToUpdate.setValues(valuesUpdate);
    } catch (err) { Browser.msgBox("Error from append without update " + err) };
}
function findChecker(checker) {
    var operTab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Oper_Tab");
    var rangeID = operTab.getRange(5, 4, operTab.getLastRow()).getValues();
    for (i in rangeID) rangeID[i] = Number(rangeID[i]);

    var indexChecker = rangeID.indexOf(Number(checker));
    var nameChecker = operTab.getRange(5 + indexChecker, 11).getValue();

    return nameChecker;
}

function getDateFromRecievedLists(link) {
    try {
        var recievedListsLink = "https://docs.google.com/spreadsheets/d/1SxVEjLlU_cRxkiXLFH_QHfk1RqAIXLuas9Rm7lu9I3o/edit#gid=1011574939"
        var curListId = SpreadsheetApp.openByUrl(link).getId();
        var recievedListSheets = SpreadsheetApp.openByUrl(recievedListsLink).getSheets();
        for (var sheet = 0; sheet < 2; sheet++) {
            var countRows = recievedListSheets[sheet].getLastRow();
            for (var rows = 0; rows < countRows; rows++) {
                var id = recievedListSheets[sheet].getRange(rows + 2, 5).getValue();
                try {
                    if (id.indexOf(curListId) != -1) {
                        var datet = recievedListSheets[sheet].getRange(rows + 2, 1).getValue().split('.');
                        return (datet[0] + '/' + datet[1] + '/' + datet[2]);
                    }
                } catch (err) { };
            }
        }
        return null;
    }
    catch (err) { Browser.msgBox(err) }
}
function countCheckers(ID, qcDateValues, qcValues, checkedByValues, qcValuesbgColors, currentDate,
    arrCheckers, arrChekerStrikes, arrChekerStrikesString, arrComments) {
    try {
        // Browser.msgBox("Count checkers");
        var rowsCount = checkedByValues.length;

        var currentRowIndex = 0;
        while (currentRowIndex < rowsCount) {
            var currentCheckedBy = checkedByValues[currentRowIndex][0];
            var curCountV = 0;
            var curStrikes = 0;
            var curStrikesString = "";
            var currentQC = false;
            while (currentRowIndex < rowsCount) {
                if (checkedByValues[currentRowIndex][0] != currentCheckedBy) {
                    break;
                }


                if (compareDates(new Date(currentDate), new Date(qcDateValues[currentRowIndex][0])) == false) {
                    currentRowIndex++; continue;
                }
                if (qcValues[currentRowIndex][0].indexOf(ID) == -1) { currentRowIndex++; continue; }
                else {
                    currentQC = true;
                    if (qcValuesbgColors[currentRowIndex][0] == "#ff0000") {
                        curStrikes++;
                        curStrikesString += qcValues[currentRowIndex][0].toString().replace(ID + " ", "") + "; ";
                    }

                    else  {
                        curCountV++;
                    }
                }
                currentRowIndex++;
            }
            if (currentQC == true) {
                arrComments.push(curCountV);
                arrCheckers.push(currentCheckedBy);
                arrChekerStrikes.push(curStrikes);
                arrChekerStrikesString.push(curStrikesString)
            }
        }
    }
    catch (err) { Browser.msgBox(err) }
}
function compareDates(date1, date2) {
    if (date1.getMonth() == date2.getMonth() && date1.getDate() == date2.getDate() && date1.getYear() == date2.getYear())
        return true;
    return false;
}


function sortCheckers(qcDateValues, qcValues, checkedByValues, qcValuesbgColors) {
    //Browser.msgBox("Sort checkers");
    var rowsCount = checkedByValues.length;
    for (var i = 0; i < rowsCount; i++) {
        var minIndex = i;
        for (var j = i + 1; j < rowsCount; j++) {
            if (checkedByValues[j][0] < checkedByValues[minIndex][0])
                minIndex = j;
        }
        if (minIndex != i) {
            swapRows(qcDateValues, minIndex, i, 0, 0);
            swapRows(qcValues, minIndex, i, 0, 0);
            swapRows(checkedByValues, minIndex, i, 0, 0);
            swapRows(qcValuesbgColors, minIndex, i, 0, 0);
        }
    }
}
function swapRows(mas, indexRow1, indexRow2, indexColumn1, indexColumn2) {
    var temp = mas[indexRow1][indexColumn1];
    mas[indexRow1][indexColumn1] = mas[indexRow2][indexColumn2];
    mas[indexRow2][indexColumn2] = temp;
}
