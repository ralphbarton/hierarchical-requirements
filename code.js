function showMessageBox() {
    Browser.msgBox('Hello World');
    changeColumns();
}


function logProductInfo() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
	Logger.log('Product name: ' + data[i][0]);
	Logger.log('Product number: ' + data[i][1]);
    }
}


var fontSizes = [18,14,12,10,8];

function changeColumns() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();

    // Let's assume a max depth of 5
    // i.e. 1.1.1.2.1

    //REMEMBER - items in Data array are 0-indexed.
    // row/column ranges are 1-indexed.


    
    for (var i = 0; i < data.length; i++) {

	var rowIndex = i+1;
	var specifiedCol = data[i][0];
	var itemTitle = data[i][1] || data[i][2] || data[i][3] || data[i][4] || data[i][5];

	// Extract text from any preceeding numbers...
	var firstLetter = itemTitle.match(/[a-zA-Z]/);
	var firstLetIndex = itemTitle.indexOf(firstLetter);
	var cleanTitle = itemTitle.substring(firstLetIndex);

	
	//oldTextCell.clearContent();
	if(specifiedCol){

	    // heartbeat
	    Logger.log(rowIndex + ": r-retrieve" + specifiedCol);

	    // clear the original text
	    // getRange(row, column, numRows, numColumns)
	    var row5cols = sheet.getRange(rowIndex, 2, 1, 5);
	    row5cols.clearContent();

	    //set font size
	    var myFontSize = fontSizes[specifiedCol-1];
	    row5cols.setFontSize(myFontSize);

	    // replace it into desired column
	    var targetCell = sheet.getRange(rowIndex, specifiedCol+1);
	    targetCell.setValue(cleanTitle);
	}
    }
}
