var numberingCounters = [0, 0, 0, 0, 0];
var prevDepth = 0;
function nextAutoNumber(depth){

    if(depth < prevDepth){

	//reset them all to a value of 0
	for(var i = prevDepth; i > depth; i--){
	    numberingCounters[i] = 0;
	}
	
    }

    numberingCounters[depth]++;
    prevDepth = depth;

    var retString = "";
    for(var i = 0; i <= depth; i++){
	retString += numberingCounters[i] + ".";
    }
    return retString + " ";
}



var fontSizes = [18, 14, 12, 10, 8];

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

	
	if(specifiedCol){


	    // Extract text from any preceeding numbers...
	    var firstLetter = itemTitle.match(/[a-zA-Z]/);
	    var firstLetIndex = itemTitle.indexOf(firstLetter);
	    var cleanTitle = itemTitle.substring(firstLetIndex);

	    var numberedTitle = nextAutoNumber(specifiedCol-1) + cleanTitle;
	    
	    // heartbeat
	    Logger.log(rowIndex + ": r-retrieve" + specifiedCol);

	    // clear the original text
	    // getRange(row, column, numRows, numColumns)
	    var row5cols = sheet.getRange(rowIndex, 2, 1, 5);
	    row5cols.clearContent();

	    //set font size
	    var myFontSize = fontSizes[specifiedCol-1];
	    row5cols.setFontSize(myFontSize);

	    //set row height, if its a level 1 heading.
	    var myHeight = specifiedCol === 1 ? 65 : 21;
	    sheet.setRowHeight(rowIndex, myHeight);
	    
	    // replace it into desired column
	    var targetCell = sheet.getRange(rowIndex, specifiedCol+1);
	    targetCell.setValue(numberedTitle);
	}
    }
}
