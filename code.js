

// I'm making this statement global
var sheet = SpreadsheetApp.getActiveSheet();


// Helper functions...
var autoNumbering = {

    numberingCounters: [0, 0, 0, 0, 0],
    prevDepth: 0,

    genNext: function(depth){

	if(depth < this.prevDepth){

	    //reset them all to a value of 0
	    for(var i = this.prevDepth; i > depth; i--){
		this.numberingCounters[i] = 0;
	    }
	    
	}

	this.numberingCounters[depth]++;
	this.prevDepth = depth;

	var retString = "";
	for(var i = 0; i <= depth; i++){
	    retString += this.numberingCounters[i] + ".";
	}
	return retString + " ";
    }
};








var fontSizes = [18, 14, 11, 10, 8];
var rowHeights = [65, 21, 21, 18, 16];

//how many colums are to the left of the top level of hierarchy?
var qtyLeftColumns = 1;
var qtyTopRows = 7;


function reformatHierarchy() {
    var data = sheet.getDataRange().getValues();

    //REMEMBER - items in Data array are 0-indexed.
    // row/column ranges are 1-indexed.
    
    // addressing cells using (R, C)
    var useIndent =    sheet.getRange(2, 16).getValue();
    var useNumbering = sheet.getRange(3, 16).getValue();    

    
    
    for (var i = 0; i < data.length; i++) {

	var rowIndex = i+1;

	//hierchy-depth of this item (row). Top-level is value 1.
	var item_HDepth = data[i][0];

	// Assumed a max depth of 5 - i.e. 1.1.1.2.1
	var itemTitle = data[i][1] || data[i][2] || data[i][3] || data[i][4] || data[i][5];

	
	if(item_HDepth){


	    // Generate new Title Text...
	    // 1. Extract text from any preceeding numbers...
	    var firstLetter = itemTitle.match(/[a-zA-Z]/);
	    var firstLetIndex = itemTitle.indexOf(firstLetter);
	    var cleanTitle = itemTitle.substring(firstLetIndex);

	    // 2. whether to 'use numbering' is an option...
	    var newTitleText = (useNumbering ? autoNumbering.genNext(item_HDepth-1) : "") + cleanTitle;
	    
	    // clear the original text
	    // getRange(row, column, numRows, numColumns)
	    var row5cols = sheet.getRange(rowIndex, qtyLeftColumns+1 , 1, 5);
	    row5cols.clearContent();

	    //set font size
	    var myFontSize = fontSizes[item_HDepth-1];
	    row5cols.setFontSize(myFontSize);

	    //set row height, if its a level 1 heading.
	    var myHeight = rowHeights[item_HDepth-1];
	    sheet.setRowHeight(rowIndex, myHeight);
	    
	    // replace it into desired column
	    var columnIndex = qtyLeftColumns + (useIndent ? item_HDepth : 1);
	    var targetCell = sheet.getRange(rowIndex, columnIndex);
	    targetCell.setValue(newTitleText);
	}
    }
}



// Composite functions

function toggleIndent(){
    var newValue = 1 - sheet.getRange(2, 16).getValue();
    sheet.getRange(2, 16).setValue(newValue);
    reformatHierarchy();
};


function toggleNumbering(){
    var newValue = 1 - sheet.getRange(3, 16).getValue();
    sheet.getRange(3, 16).setValue(newValue);
    reformatHierarchy();
};

function newRows(qty){
    var ActiveCell = sheet.getActiveCell();
    var ActiveRow = ActiveCell.getRow();
    
    // 1. add the new rows
    sheet.insertRowsAfter(ActiveRow, qty);

    // 2. format all those new rows...
    var newrows_HDepth = sheet.getRange(ActiveRow, 1).getValue();
    
    for (var i = 1; i <= qty; i++) {
	var row_i = ActiveRow + i;

	// (a) set depth number
	sheet.getRange(row_i, 1).setValue(newrows_HDepth);

	// (b) set font size (to item depth 3 setting, regardless of actual)
	sheet.getRange(row_i, qtyLeftColumns+1, 1, 5).setFontSize(fontSizes[3]);;
	
	// (c) set row height (to item depth 3 setting, regardless of actual)
	sheet.setRowHeight(row_i, rowHeights[3]);
    }
};

function newRows_four(){
    newRows(4);
}

function newRows_twelve(){
    newRows(12);
}

function deleteUnusedRows(){
    var data = sheet.getDataRange().getValues();

    var qty_rows_already_deleted = 0;
    
    for (var i = qtyTopRows; i < data.length; i++) {
	// Assumed a max depth of 5 - i.e. 1.1.1.2.1
	var itemTitle = data[i][1] || data[i][2] || data[i][3] || data[i][4] || data[i][5];

/*
	var str = "" + i + "    " + itemTitle + " = " + (!itemTitle);
	Logger.log(str); 
*/
	//no item title, delete this row
	if(!itemTitle){
	    // offset for (RC 1-based index vs Array 0-based index) AND rows already deleted.
	    sheet.deleteRow(i + 1 - qty_rows_already_deleted);
	    qty_rows_already_deleted++;
	}

//	if(i>20){break;}
	
    }
	

}
