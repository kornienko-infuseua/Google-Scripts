
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Check sheets for DB')
        .addItem('Select Date', 'createDialog')
        .addToUi();

}
function onEdit(e) { // when teamleads put 'done', replace it with 'done' 
    var range = e.range;
    var sheet = SpreadsheetApp.getActiveSheet();
    if (range.getColumn() == 8 && range.getValue() == "done") {
        sheet.getRange(range.getRow(), range.getColumn()).setValue("Done");
    }
}
function createDialog() {
    var htmlDialog = HtmlService.createHtmlOutputFromFile("scriptHTML")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(300)
        .setWidth(270);
    SpreadsheetApp.getUi().showModalDialog(htmlDialog, "Select Date");
}
var SheetController = function (dateToScript, isWholeMonth, checkRejectionRate) {
    this.dateToScript = dateToScript;
    this.isWholeMonth = isWholeMonth;
    this.checkRejectionRate = checkRejectionRate;
    this.currentDate = generateCurrentDate();
}

function useScript(dateToScript, isWholeMonth, checkRejectionRate) {
    // dateToScript = date from htmlForm in 'mm.dd.yyyy' format
    try {

        //// var dateToScript = '05.16.2018';
        // var isWholeMonth = false;
        //var checkRejectionRate = true;
        var sheetController = new SheetController(dateToScript, isWholeMonth, checkRejectionRate)

        var wholeTable = GetRangeToProcess(sheetController) //get rows from received lists with selected date OR select all rows if WholeMonth is enabled

        if (wholeTable == false) { // don't do anything if can not get sheet by name
            return;
        }
        var data = wholeTable.getValues();
        var rowsToScript = getRowsToScript(data, isWholeMonth, dateToScript)


        var dateColumn = 0, listNameColumn = 1, commentColumn = 5, amountOfLeadsColumn = 2, linkColumn = 4, statusColumn = 7, dateSciptColumn = 8, scriptColumn = 9, readColumn = 10;

        var rejSheetArr = [] //adding each list into array to append rows in rejection rate
        //adding each new/wrong email into array to append rows in List Errors (Email analysis tab)

        for (row in rowsToScript) {
            try {
                var currentRow = rowsToScript[row];
                var currrentSheet = OpenSheetByLink(data[currentRow][linkColumn]) // try to open sheet nad don't process it if can't access
                if (!currrentSheet) {
                    data[currentRow][scriptColumn] = "No Permission";
                }
                else {
                    var result = cleanTheList(currrentSheet, data[currentRow][linkColumn], data[currentRow][dateColumn]);
                    //cleanTheList: delete empty rows, clear prooflinks, make column headers lowercase, 
                    // return 'mistake' property if list doesn't contain any required column
                    if (result.res || data[currentRow][scriptColumn].toString() == "ok") {
                        data[currentRow][scriptColumn] = "done";
                        data[currentRow][dateSciptColumn] = sheetController.currentDate;
                        if (data[currentRow][commentColumn].toString() == "unChecked") {
                            data[currentRow][commentColumn] = "";
                        }
                        if (checkRejectionRate) { // Get statistics about rejection rate
                            //var rejRes = getRejectionRate(result, currrentSheet, data[currentRow][listName], data[currentRow][dateColumn], data[currentRow][linkColumn]);
                            var rejRes = getRejectionRateNew(result, currrentSheet, data[currentRow][listNameColumn], data[currentRow][dateColumn], data[currentRow][linkColumn]);
                            if (rejRes) {
                                var re = CreateRejRow(rejRes, sheetController.currentDate, data[currentRow][dateColumn], data[currentRow][listNameColumn], data[currentRow][linkColumn])
                                rejSheetArr.push(re)
                            }
                        }

                    }
                    else if (result == false) data[currentRow][commentColumn] = isWholeMonth ? "UNChecked" : "unChecked"; // put unckecked if there is no checked leads (all ov_comments are empty)

                    else {
                        data[currentRow][scriptColumn] = "missing " + result.mistakes;
                    }
                    /* if (result.newEmails) {
                        data[currentRow][dateSciptColumn] = result.newEmails;
                    } */

                    if (isWholeMonth) {
                        wholeTable.setValues(data)
                    }

                }


            }
            catch (err) { Browser.msgBox("in cycle + " + err) }
        }
        if (result && result.NewEmails && result.NewEmails.length > 0) InputNewEmails(result.NewEmails); // input wrong/new emails into List Errors spreadsheet, Email analysis tag
        if (!isWholeMonth) {
            wholeTable.setValues(data)
        }
        if (checkRejectionRate) {
            InputRejRate(rejSheetArr)
        }

    } catch (err) { Browser.msgBox("Error in " + err) }
    //if (!isWholeMonth) Browser.msgBox("Script is Done");
}
function CreateRejRow(rejCount, dateScript, dateList, listName, link) {
    var count = rejCount.countChecked;
    function GetPercent(val) {
        return ((val / count * 100).toFixed(2)) + "%";
    }

    return [
        dateScript, dateList, listName, link,
        GetPercent(rejCount.titleGreen), rejCount.titleGreen,
        GetPercent(rejCount.titleYellow), rejCount.titleYellow,

        GetPercent(rejCount.countryGreen), rejCount.countryGreen,
        GetPercent(rejCount.countryYellow), rejCount.countryYellow,

        GetPercent(rejCount.industryGreen), rejCount.industryGreen,
        GetPercent(rejCount.industryYellow), rejCount.industryYellow,

        GetPercent(rejCount.employeesGreen), rejCount.employeesGreen,
        GetPercent(rejCount.employeesYellow), rejCount.employeesYellow,

        GetPercent(rejCount.revenueGreen), rejCount.revenueGreen,
        GetPercent(rejCount.revenueYellow), rejCount.revenueYellow,

        GetPercent(rejCount.nacsup), rejCount.nacsup,
        GetPercent(rejCount.qTitle), rejCount.qTitle,
        GetPercent(rejCount.qCompany), rejCount.qCompany,
        GetPercent(rejCount.qOther), rejCount.qOther,
        rejCount.countChecked]
}
function InputRejRate(rejSheetArr) {
    try {
        var rejSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/13vyGSPKDGfFJiw_tc8ahnGxM2T6GiIXpfZLjZJm6UNI/edit#gid=0").getSheetByName('New lists');
        if (rejSheetArr.length > 0) rejSheet.getRange(rejSheet.getLastRow() + 1, 1, rejSheetArr.length, rejSheetArr[0].length).setValues(rejSheetArr);

    } catch (err) { Browser.msgBox("Can not open Rejection sheet " + err) }

}

function InputNewEmails(NewEmails) {

    try {
        var wrongEmailSheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1w_rH9yTNw4CTLB4nhT2H2qlsTMtvjeZgP3b1xIfgy60/edit?ts=5947e50d").getSheetByName('Email analysis');
        var Arr = [];

        for (var i = 0; i < NewEmails.length; i++) {
            Arr.push([NewEmails[i].Date, NewEmails[i].OldEmail ? NewEmails[i].OldEmail : "", NewEmails[i].NewEmail ? NewEmails[i].NewEmail : "", NewEmails[i].Link, NewEmails[i].Prooflink])

        }
        wrongEmailSheet.getRange(wrongEmailSheet.getLastRow() + 1, 1, Arr.length, Arr[0].length).setValues(Arr);
    } catch (err) { Browser.msgBox("Can not Open List Errors " + err) }

}
function GetRangeToProcess(sheetController) {
    try {
        var recl = SpreadsheetApp.getActiveSpreadsheet();
        if (sheetController.isWholeMonth) {
            var sheetToScript = SpreadsheetApp.getActiveSheet();
        }
        else {
            var nameOfSheet = getNameOfSheet(sheetController.dateToScript);

            var sheetToScript = recl.getSheetByName(nameOfSheet);
            if (sheetToScript) {
                var lastColumn = sheetToScript.getLastColumn();
                if (!sheetController.isWholeMonth) {
                    var dateRange = sheetToScript.getRange(1, 1, sheetToScript.getLastRow());
                    var dateValues = dateRange.getValues();
                    var stopPosition = 0;
                    var startPosition = 0;
                    while (stopPosition < dateValues.length) {
                        if (sheetController.dateToScript == dateValues[stopPosition]) {
                            startPosition = stopPosition;
                            while (++stopPosition < dateValues.length) {
                                if (sheetController.dateToScript != dateValues[stopPosition])
                                    break;
                            }
                            break;
                        }
                        stopPosition++;
                    }
                    stopPosition++;
                }
                else {
                    startPosition = 1; stopPosition = sheetToScript.getLastRow();
                    Browser.msgBox("Whole month");
                }
                var countRows = stopPosition - startPosition - 1;
                if (countRows <= 0) { Browser.msgBox("No list with selected date"); return 0; }

                var wholeTable = sheetToScript.getRange(startPosition + 1, 1, countRows, lastColumn);

                return wholeTable;
            }
            else {

                Browser.msgBox(nameOfSheet + " is missing");
                return false;
            }
        }
    } catch (err) {
        Browser.msgBox("Can not get rows from received lists" + err)
    }
}
function getMissingColumns(sheet, columnNames, newColumnNames) {
    try {
        //var sheet = SpreadsheetApp.openByUrl(link).getSheets()[0];
        var range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
        var values = range.getValues();

        for (i in values[0])
            values[0][i] = replaceSpaces(values[0][i])
        for (column in columnNames) {
            var index = values[0].indexOf(columnNames[column]); //UPDATE 28.11   !!!!!!!!!!!!!!!!!!!!!!!
            if (index != -1) {
                newColumnNames[columnNames[column]] = index;
            }
            else newColumnNames[columnNames[column]] = columnNames[column];

        }

        var mistakes_test = "";
        for (column in columnNames) {
            if (typeof (newColumnNames[columnNames[column]]) != 'number')
                mistakes_test += newColumnNames[columnNames[column]] + "; "
        }
        range.setValues(values);

    } catch (err) { Browser.msgBox(err) }

    return mistakes_test;
}
function cleanTheList(currentList, linkList, listdate) {
    try {
        var result = {}
        result.NewEmails = []
        var lastColumn = currentList.getLastColumn();
        var lastRow = currentList.getLastRow();
        result.LastColumn = lastColumn;
        result.LastRow = lastRow;
        var range = currentList.getRange(1, 1, lastRow, lastColumn);
        var values = range.getValues();
        var maxRows = currentList.getMaxRows();
        var countRows = maxRows - lastRow;

        if (countRows != 0) currentList.deleteRows(lastRow + 1, maxRows - lastRow); //delete empty rows;
        var newColumnNames = []
        var columnNames = ["first_name", "last_name", "company", "title", "email", "address", "city", "state", "zip", "country", "phone", "prooflink", "employees", "employees_prooflink", "revenue", "revenue_prooflink", "ov_comment", "industry"];
        var mistakes = getMissingColumns(currentList, columnNames, newColumnNames)
        var bgColors = range.getBackgrounds();
        var weights = range.getFontWeights();
        var fontColors = range.getFontColors();

        newEmailRows = []
        var emailColumn = newColumnNames['email']
        var prooflink = newColumnNames['prooflink']
        var employees_prooflink = newColumnNames['employees_prooflink']
        var revenue_prooflink = newColumnNames['revenue_prooflink']

        for (var i = 1; i < lastRow; i++) {

            if (typeof prooflink == 'number') {
                var link_str = values[i][prooflink].toString();
                if (link_str.indexOf('linkedin') != -1)
                    values[i][prooflink] = link_str.split('?')[0];
            }

            if (typeof employees_prooflink == 'number') {
                link_str = values[i][employees_prooflink].toString();
                if (link_str.indexOf('yahoo') != -1 || link_str.indexOf('linkedin') != -1)
                    values[i][employees_prooflink] = link_str.split('?')[0];
            }

            if (typeof revenue_prooflink == 'number') {
                link_str = values[i][revenue_prooflink].toString();
                if (link_str.indexOf('yahoo') != -1 || link_str.indexOf('linkedin') != -1)
                    values[i][revenue_prooflink] = link_str.split('?')[0];
            }
            //check email

            if (typeof emailColumn == 'number' && bgColors[0][emailColumn] != "#f5bfb3" && fontColors[i][emailColumn] == "#ff0000" && weights[i][emailColumn] == "bold") {
                newEmailRows.push(i)
            }
        }


        try {
            if (newEmailRows.length > 0) {
                var notes = currentList.getRange(1, emailColumn + 1, lastRow).getNotes();
                for (var i = 0; i < newEmailRows.length; i++) {
                    var newEmail = {};
                    var curRow = newEmailRows[i];
                    if (bgColors[curRow][emailColumn] == "#ffff00") {
                        newEmail.NewEmail = values[curRow][emailColumn]
                        if (notes[curRow][0]) {
                            newEmail.OldEmail = notes[curRow][0]
                        }
                    }
                    else {
                        newEmail.OldEmail = values[curRow][emailColumn]
                    }
                    newEmail.Prooflink = values[curRow][prooflink]
                    newEmail.Link = linkList;
                    newEmail.Date = listdate;
                    result.NewEmails.push(newEmail)
                }
                setBackground(currentList.getRange(1, emailColumn + 1), "#f5bfb3")
            }
        }
        catch (err) { Browser.msgBox("4.0 " + err) }

        /*  if (newEmailRows) {
             result.newEmails = "New emails found: " + newEmailRows
         } */

        if (mistakes) {
            result.mistakes = mistakes;
            return result;
        }

        if (!anyChecked(values, newColumnNames['ov_comment'])) {
            return false;
        }
        /*   var isUnCheckedList = true; //check if list is fully uncheked (no color coding for: title,phone, prooflink)
          for (var i = 1; i < lastRow; i++) {
              if (bgColors[i][newColumnNames['prooflink']] != "#ffffff") { isUnCheckedList = false; break; }
  
              else if (bgColors[i][newColumnNames['phone']] != "#ffffff") { isUnCheckedList = false; break; }
  
              else if (bgColors[i][newColumnNames['title']] != "#ffffff") { isUnCheckedList = false; break; }
          } */
        range.setValues(values);
        /* if (isUnCheckedList) {
            return false;
        } */
        result.ColumnNames = newColumnNames;
        result.res = true;
        result.range = range;
        result.values = values;

        return result;

    } catch (err) { Browser.msgBox("exception: " + err) }
}

function setBackground(cell, colour) {
    cell.setBackground(colour)
}
function anyChecked(values, ovCommentColumn) {
    if (ovCommentColumn >= 0) {
        for (var i = 0; i < values.length; i++) {
            if (values[i][ovCommentColumn].toString() != "" || values[i][ovCommentColumn].toString() != "N/A") {
                return true;
            }
            return false;
        }
    }
    return false;
}
function ifListChecked(link) {
    try {
        var curSheet = SpreadsheetApp.openByUrl(link).getSheets()[0];
        var lastRow = curSheet.getLastRow();
        if (curSheet.getMaxRows() - lastRow > 0)
            curSheet.deleteRows(lastRow + 1, curSheet.getMaxRows() - lastRow); //delete empty rows;
        var columns = curSheet.getRange(1, 1, 1, curSheet.getLastColumn()).getValues();
        var ovCommentColumn = columns[0].indexOf("ov_comment") + 1;
        var pvCommentColumn = columns[0].indexOf("pv_comment") + 1;
        if (!ovCommentColumn) { return false; }
        var ovCommentRows = curSheet.getRange(2, ovCommentColumn, lastRow - 1).getValues();
        if (pvCommentColumn) var pvCommentRows = curSheet.getRange(2, pvCommentColumn, lastRow - 1).getValues();
        for (var row = 0; row < ovCommentRows.length; row++) {
            if (ovCommentRows[row][0] == "") return false;

            if (ovCommentRows[row][0].indexOf('Y') == 0) {
                if (!pvCommentColumn)
                    return false;
                var pvcom = pvCommentRows[row][0].toString().toLowerCase();
                if (pvcom == "" || pvcom == "n/a")
                    return false;
            }
        }
        return true;
    } catch (err) { return "No permissions"; }
}

function OpenSheetByLink(link) {
    try {
        var sheet = SpreadsheetApp.openByUrl(link).getSheets()[0];;
        return sheet;
    } catch (err) {
        return false;
    }
}
function getRowsToScript(data, isWholeMonth, dateToScript) {
    var masRows = []
    var dateColumn = 0, commentColumn = 5, amountOfLeadsColumn = 2, linkColumn = 4, statusColumn = 7, dateSciptColumn = 8, scriptColumn = 9, readColumn = 10;
    if (!isWholeMonth) {

        for (var i = 0; i < data.length; i++) {


            if (data[i][dateColumn] != dateToScript) continue;
            if (data[i][scriptColumn] == "done" || data[i][linkColumn] == 0) continue;
            var status = data[i][statusColumn].toString().toLowerCase();
            var comment = data[i][commentColumn].toString().toLowerCase();

            if (data[i][amountOfLeadsColumn] > 50) {
                if ((status != "done" && comment == "") || (comment != "platform" && status == "")) {
                    if (!ifListChecked(data[i][linkColumn])) continue;
                }
            }

            if (data[i][commentColumn].toString().toLowerCase() == "no db") continue;

            masRows.push(i);

        }
    }
    else {

        for (var i = 0; i < data.length; i++) {
            if (data[i][commentColumn] == "UNChecked") continue;
            if (data[i][scriptColumn] != "" || data[i][linkColumn] == 0) continue;
            if (data[i][commentColumn].toString().toLowerCase() == "no db") continue;
            masRows.push(i);
        }
    }
    //Browser.msgBox("Rows to script: " + masRows.length)
    //Browser.msgBox(masRows.length)
    return masRows
}
function getRejectionRateNew(result, curSheet, sheetName, date, link) {
    try {
        var OV_comments = ["y1: linkedin/company website", "y2: pl summary", "y3: facebook", "y4: suspicious linkedin", "y5: 3rd party prooflink", "n1: nwc", "n2: out of business/bad data", "n/a: pv tool", "n/a: title/pl summary", "n/a: industry", "n/a: emp. size", "n/a: revenue", "n/a: probably nwc", "n/a: country/geo", "n/a: nac/sup", "n/a: nac", "nac", "n/a: prooflink", "n/a: wrong email/general domain", "n/a: other", "q1: questionable title", "q2: questionable company", "q3: other", "n/a: country", "n/a: geo"];
        var ovCommentColumn = result.ColumnNames['ov_comment']
        var titleColumn = result.ColumnNames['title']
        var countryColumn = result.ColumnNames['country']
        var employeesColumn = result.ColumnNames['employees']
        var revenueColumn = result.ColumnNames['revenue']
        var industryColumn = result.ColumnNames['industry']
        var companyColumn = result.ColumnNames['company']
        var values = result.values;

        if (!ovCommentColumn) {
            return false;
        }
        var countChecked = 0;
        var rejCount = {};
        rejCount['titleGreen'] = 0;
        rejCount['titleYellow'] = 0;
        rejCount['industryGreen'] = 0;
        rejCount['industryYellow'] = 0;
        rejCount['countryGreen'] = 0;
        rejCount['countryYellow'] = 0;
        rejCount['employeesGreen'] = 0;
        rejCount['employeesYellow'] = 0;
        rejCount['revenueGreen'] = 0;
        rejCount['revenueYellow'] = 0;
        rejCount['nacsup'] = 0;
        rejCount['qTitle'] = 0;
        rejCount['qCompany'] = 0;
        rejCount['qOther'] = 0;
        var valuesColors = result.range.getBackgrounds();
        for (var i = 1; i < values.length; i++) {
            var curComment = values[i][ovCommentColumn].toString().toLowerCase();
            if (OV_comments.indexOf(curComment) != -1) {
                countChecked++;
                if (curComment.search('n/a: title') != -1) {
                    GetRejCounts(rejCount, valuesColors, i, titleColumn, 'title')
                } else if (curComment.search('n/a: industry') != -1) {
                    GetRejCounts(rejCount, valuesColors, i, industryColumn, 'industry')
                }
                else if (curComment.search('n/a: emp') != -1) {
                    GetRejCounts(rejCount, valuesColors, i, employeesColumn, 'employees')
                }
                else if (curComment.search('n/a: rev') != -1) {
                    GetRejCounts(rejCount, valuesColors, i, revenueColumn, 'revenue')
                }
                else if (curComment.search('country') != -1 || curComment.search('geo') != -1) {
                    GetRejCounts(rejCount, valuesColors, i, countryColumn, 'country')
                }
                else if (curComment.search('nac') != -1 || curComment.search('sup') != -1) {
                    rejCount['nacsup']++
                }
                else if (curComment.indexOf('q1') == 0) {
                    rejCount['qTitle']++
                } else if (curComment.indexOf('q2') == 0) {
                    rejCount['qCompany']++
                }
                else if (curComment.indexOf('q3') == 0) {
                    rejCount['qOther']++
                }
            }
        }
        rejCount.countChecked = countChecked;
        return rejCount;
    }
    catch (err) {
        Browser.msgBox("Err in rej rate" + err)
    }
}

function GetRejCounts(rejCount, valuesColors, curIndex, rejColumnName, rejReason) {
    if (valuesColors[curIndex][rejColumnName] == "#93c47d") {
        rejCount[rejReason + 'Green']++;
    }
    else {
        rejCount[rejReason + 'Yellow']++;
    }
}

/* function getCheckedLeads(curSheet, lastRow, OV_CommentColumn) {
    var OV_comments = ["Y1: linkedin/company website", "Y2: PL Summary", "Y3: Facebook", "Y4: Suspicious Linkedin", "Y5: 3rd Party Prooflink", "N1: NWC", "N2: Out of Business/Bad data", "N/A: PV Tool", "N/A: Title/PL Summary", "N/A: Industry", "N/A: Emp. Size", "N/A: Revenue", "N/A: Probably NWC",
        "N/A: Country/GEO", "N/A: NAC/SUP", "N/A: NAC", "NAC", "N/A: Prooflink", "N/A: Wrong email/General domain", "N/A: Other", "Q1: Questionable Title", "Q2: Questionable Company", "Q3: Other", "N/A: Country", "N/A: GEO"];
    // var columns = curSheet.getRange(1, 1, 1, curSheet.getLastColumn()).getValues();
    //var OV_CommentColumn = columns[0].indexOf("ov_comment") + 1;
    if (!(OV_CommentColumn >= 0)) return -1;
    var OV_comments_all = curSheet.getRange(2, OV_CommentColumn + 1, lastRow).getValues();

    var countCheckedLeads = 0;
    OV_comments_all.forEach(function (cell) {
        if (OV_comments.indexOf(cell.toString()) != -1)
            countCheckedLeads++;
    });

    return countCheckedLeads;
} */

function getNameOfSheet(dateSelected) { //not work yet

    var nameOfSheet = "";
    var months = ["Jan", "Feb", "March", "April", "May", "June", "July", "Aug", "Sep", "Oct", "Nov", "Dec"];
    var dateArr = dateSelected.split('.');

    var year = "" + dateArr[2];
    year = year.replace("20", "");
    var a = months[parseInt(dateArr[0]) - 1]
    return (a + "_" + year);
}
function replaceSpaces(text) {
    text = text.toString().toLowerCase();
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

