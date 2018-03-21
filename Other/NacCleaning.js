function onOpen() {
         SpreadsheetApp.getUi()
        .createMenu('nac Script')
        .addItem('Go', 'deleteSymbols')
        .addItem('sort', 'sortSymbols')
        .addToUi();
}
function sortSymbols()
{
    var symSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Symbols");
    
    for (var c = 1; c<symSheet.getLastColumn(); c++)
    {
        var range = symSheet.getRange(2,c,symSheet.getLastRow());
        var newValues = testSort(symSheet,range )
        range.setValues(newValues)
    }
     
   
    
    //for (var column=1; column< symSheet.getLastColumn(); column ++)
   // {
    //    var range = symSheet.getRange(2,column,symSheet.getLastRow())
   //     var values = range.getValues()
    //    Browser.msgBox(values.length)
   // }
    
    //var range = symSheet.getRange(2,1,symSheet.getLastRow(),symSheet.getLastColumn());
   // var values = range.getValues();
   
}
function testSort(symSheet,range)
{
    //Browser.msgBox("testSort")
   
    var values = range.getValues();
   // var column = []
    for (var r=0; r<values.length; r++)
    {
        if (values[r][0]=="") return values;
       for (var r2=r+1; r2<values.length;r2 ++)
       {
        
        if (values[r2][0] == "") {
        //Browser.msgBox("break");
        break; }
           // Browser.msgBox(values[r][0] + "||" + values[r2][0] )
           //Browser.msgBox(values[r2][0].indexOf(values[r][0]))
           
           if (values[r2][0].indexOf(values[r][0])!=-1 )
           {
               //Browser.msgBox("indexOF")
               var temp = values[r2][0];
               values[r2][0] = values[r][0];
               values[r][0] = temp;
           }
       }
     }
     
     
}
function deleteSymbols() {
    var allReplaces = getSymbols();
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NAC TEST")
   
   // var allSymbols = [];
   
   // Browser.msgBox(allReplaces.length);
    deleteFromSheet(sheet,allReplaces);

    
    Browser.msgBox('Finished')
}

function deleteFromSheet(sheet, allReplaces)
{
     var lastRow = sheet.getLastRow()
    var range = sheet.getRange(1, 1, lastRow);
    var valuesNacs = range.getValues();
 try {
    //var arrLtd = [', Ltd.','. Ltd.',' ltd.',', Ltd',   ',Ltd', ' Ltd', ', Limited.',', Limited', ',Limited.', ' Limited.', ' Limited', ' LIMITED',   ' L TD.', ' L TD', ' ltd.', ' ltd', ' Limit.', ' Limit', ',LTD.', ', LTD.',' LTD.', ' LTD', ' L.T.D.', ' L.T.D', ', L.T.D.', ', L.T.D', 'L.T.D.', 'L.T.D'];
  
    for (var i=0; i<valuesNacs.length;i++)
    {
        var str = valuesNacs[i][0].toString();
        for (j in allReplaces)
        {
        

            for ( s in allReplaces[j])
            {
            var symbol = allReplaces[j][s]
            var length = symbol.length;
            
             var indexOf = str.indexOf(symbol);
                if (indexOf >=0) 
                {
                    if (indexOf+length>=str.length || str[indexOf + length] == " " || str[indexOf + length] == ",")
                    {
                  
                         valuesNacs[i][0] = str.replace(symbol, "");
                         str =  valuesNacs[i][0];
                         break;
                     }
                }
                
               // Browser.msgBox("break");
            }
        
        } 
        
    }
    range.setValues(valuesNacs)
    } catch (err) {Browser.msgBox(err) }
   
}
function getSymbols()
{
    var allReplaces = [];
    var sheetSymbols = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Symbols");
    for (var column=1; column<= sheetSymbols.getLastColumn();column++)
    {
        var values = sheetSymbols.getRange(2, column,sheetSymbols.getLastRow()).getValues();
        var mas= []
        for (var r = 0; r<values.length;r ++)
        {
            if (values[r][0]!="")
            mas[r] = values[r][0];
        }
        allReplaces.push(mas);
    }
   // Browser.msgBox("done")
    return allReplaces;
}
function replaceStr(str, oldStr, newStr)
{
     return str.replace(oldStr, newStr)
}
