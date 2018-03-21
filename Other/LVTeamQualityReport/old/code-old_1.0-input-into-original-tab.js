function onOpen() {
	SpreadsheetApp.getUi()
        .createMenu('new Report Script')
        .addItem('Select Date', 'createDialogNew')
        .addToUi();
}

function createDialogNew() {
	var htmlDialog = HtmlService.createHtmlOutputFromFile("newScriptDialog.html")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(300)
        .setWidth(270);
	SpreadsheetApp.getUi().showModalDialog(htmlDialog, "Select Date");
}
function doReport(listLink, listType, dayToScript) {
	try {
		Browser.msgBox("Report Started");
		var currentSheet = SpreadsheetApp.getActiveSheet();
		var operator = currentSheet.getRange(1, 3).getValue();
		if (!operator) { Browser.msgBox("No operator"); return false; }
		var QCSchedule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QC_Schedule");
		var rangeQCID = QCSchedule.getRange(25, 1, QCSchedule.getLastRow(), 3).getValues();

		var QC_ID;
		for (var i = 0; i < rangeQCID.length; i++)
			if (rangeQCID[i][0] == operator.toString()) { QC_ID = rangeQCID[i][2]; break; }
		if (!QC_ID) { Browser.msgBox("NO QC OPERATOR"); return 0; }
		var result = reportList(currentSheet, listType, dayToScript, listLink, QC_ID);
	}
	catch (err) { Browser.msgBox(err) }
}

function reportList(currentSheet, listType, dayToScript, listLink, ID) {
	try {
		var qcDateColumn = 1, qcColumn = 2, checkedByColumn = 3;
		var listToReport = SpreadsheetApp.openByUrl(listLink).getSheets()[0];
		var rowsCount = listToReport.getLastRow() -1;
		
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
		if (!arrCheckers) return;
		var dateFromRecievedList = getDateFromRecievedLists(listLink)
		if (!dateFromRecievedList) { Browser.msgBox("Missing in Recieved Lists"); return; }
		var temp = dateFromRecievedList.split('.');
		var outPutDate = new Date(temp[0] + '/' + temp[1] + '/' + temp[2]);
	}
	catch (err) { Browser.msgBox("Error from reportList " + err) }
	appendRows(currentSheet, listLink, listType,
        arrCheckers, arrChekerStrikes, arrChekerStrikesString, arrComments, outPutDate, currentDate)
	Browser.msgBox("Report finished");
}

function appendRows(currentSheet, listLink, listType,
    arrCheckers, arrChekerStrikes, arrChekerStrikesString, arrComments, outPutDate, currentDate) {
	try {
		var r = currentSheet.getRange(6, 1, currentSheet.getLastRow()).getValues();
		var links = currentSheet.getRange(6, 5, currentSheet.getLastRow()).getValues();
		var countNotEmptyRows = 0;

		var curListId = SpreadsheetApp.openByUrl(listLink).getId();
		while (countNotEmptyRows < r.length) {
			if (links[countNotEmptyRows][0].indexOf(curListId) != -1) {
				if (compareDates(new Date(currentDate), currentSheet.getRange(countNotEmptyRows + 6, 1).getValue())) {
					var checkerName = currentSheet.getRange(6 + countNotEmptyRows, 3).getValue();
					checker = checkerName.split(' ')[0];
					var index = arrCheckers.indexOf(Number(checker));
					if (index > -1) {
						var updateCheckerRange = currentSheet.getRange(6 + countNotEmptyRows, 8, 2, 3);
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
		var dateFromRL = getDateFromRecievedLists(listLink);
		var operTab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Oper_Tab");
		var rangeID = operTab.getRange(5, 4, operTab.getLastRow()).getValues();
		for (i in rangeID) rangeID[i] = Number(rangeID[i]);
		var rangeToUpdate = currentSheet.getRange(countNotEmptyRows + 6, 1, arrCheckers.length, 10);
		var valuesUpdate = rangeToUpdate.getValues();
		var count = 0;
		for (var i = 0; i < arrCheckers.length; i++) {
			if (!arrCheckers[i]) continue;
			var indexChecker = rangeID.indexOf(Number(arrCheckers[i]));
			if (indexChecker == -1) continue;
			var nameChecker = operTab.getRange(5 + indexChecker, 11).getValue();
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
	} catch (err) { Browser.msgBox("Error from append " + err) };
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
		var recievedListsLink = "https://docs.google.com/spreadsheets/d/1KlyQ6YpLAGPh_pQL3Ovw24sGZi0ZamJSoFEMfXvLTWQ/edit#gid=538113528"
		var curListId = SpreadsheetApp.openByUrl(link).getId();
		var recievedListSheets = SpreadsheetApp.openByUrl(recievedListsLink).getSheets();
		for (var sheet = 0; sheet < recievedListSheets.length; sheet++) {
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

					else if (qcValues[currentRowIndex][0] == ID + " v") {
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
