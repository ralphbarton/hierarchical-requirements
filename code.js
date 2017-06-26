

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
var qtyLeftColumns = 2;
var qtyTopRows = 7;
var sheetGlobalsColumn = 18;


function getItemTitle(row_data) {
    // row_data[] is a zer-indexed array
    const ofs = qtyLeftColumns - 1;
    // Assumed a max depth of 5 - i.e. 1.1.1.2.1
    return row_data[ofs + 1] || row_data[ofs + 2] || row_data[ofs + 3] || row_data[ofs + 4] || row_data[ofs + 5];
}



function reformatHierarchy(limit) {
    //before even starting delete unused rows
    deleteUnusedRows();

    var data = sheet.getDataRange().getValues();

    //REMEMBER - items in Data array are 0-indexed.
    // row/column ranges are 1-indexed.
    
    // addressing cells using (R, C)
    var uidCount = sheet.getRange(1, sheetGlobalsColumn).getValue();
    var useIndent =    sheet.getRange(2, sheetGlobalsColumn).getValue();
    var useNumbering = sheet.getRange(3, sheetGlobalsColumn).getValue();
    var HDepth_bold = sheet.getRange(5, sheetGlobalsColumn).getValue();
   

    var range = sheet.getActiveRange();
    var topActiveRow = range.getRowIndex();
    var activeHeight = range.getHeight();
    
    
    
    
    for (var i = 0; i < data.length; i++) {

	var rowIndex = i+1;

	//hierchy-depth of this item (row). Top-level is value 1.
	var item_HDepth = data[i][0];
	
	var itemTitle = getItemTitle(data[i]);

	// if multiple rows selected AND i is outside range
	var thisRowWithin = (rowIndex >= topActiveRow) && (rowIndex < (topActiveRow+activeHeight));
	var skipRow = (activeHeight > 1) && (!thisRowWithin);
	
	if((item_HDepth)&&(typeof(item_HDepth)==="number")&&(itemTitle)&&(!skipRow)){


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

	    // (A) set font size (based upon 'item_HDepth')
	    var myFontSize = fontSizes[item_HDepth-1];
	    Logger.log(item_HDepth + "  /  " + myFontSize);
	    row5cols.setFontSize(myFontSize);

	    // (B) emboldenment (based upon 'item_HDepth'). This includes colour
	    const no_bold = HDepth_bold === 'x';
	    row5cols.setFontWeight(item_HDepth === HDepth_bold ? 'bold' : 'normal');
	    row5cols.setFontColor((item_HDepth === HDepth_bold) || no_bold ? 'black' : 'grey');

	    // (C) set row height (based upon 'item_HDepth')
	    var myHeight = rowHeights[item_HDepth-1];
	    sheet.setRowHeight(rowIndex, myHeight);
	    
	    // (D) write the text into desired column
	    var HDepthColumnIndex = qtyLeftColumns + (useIndent ? item_HDepth : 1);
	    sheet.getRange(rowIndex, HDepthColumnIndex).setValue(newTitleText);

	    // (E) give it a UID, if uid is missing, and increment...
	    if(!data[i][1]){// missing uid
		sheet.getRange(rowIndex, 2).setValue(uidCount);
		uidCount++;
	    }
	}
    }
    
    // Now that loop has finished, rewrite the final UID
    sheet.getRange(1, sheetGlobalsColumn).setValue(uidCount);
    
    
}



// Composite functions

function toggleIndent(){
    var newValue = 1 - sheet.getRange(2, sheetGlobalsColumn).getValue();
    sheet.getRange(2, sheetGlobalsColumn).setValue(newValue);
    reformatHierarchy();
};


function toggleNumbering(){
    var newValue = 1 - sheet.getRange(3, sheetGlobalsColumn).getValue();
    sheet.getRange(3, sheetGlobalsColumn).setValue(newValue);
    reformatHierarchy();
};

function newRows(qty){
    var ActiveCell = sheet.getActiveCell();
    var ActiveRow = ActiveCell.getRow();
    
    // 1. add the new rows
    sheet.insertRowsAfter(ActiveRow, qty);

    // 2. format all those new rows...
    const targetRow_HDepth = sheet.getRange(ActiveRow, 1).getValue();
    const newrows_HDepth = Math.min(targetRow_HDepth+1, 5);
    
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
	var itemTitle = getItemTitle(data[i]);

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


function hideRowsPastDepth(){
    var data = sheet.getDataRange().getValues();

    Logger.log("data.length = " + data.length); 

    var HDepthLim = sheet.getRange(4, sheetGlobalsColumn).getValue();
    
    for (var i = qtyTopRows; i < data.length; i++) {

	//hierchy-depth of this item (row). Top-level is value 1.
	var item_HDepth = data[i][0];

	if(item_HDepth > HDepthLim){
	    sheet.hideRows(i+1);
	}else{
	    sheet.showRows(i+1);
	}
	
    }
    
};


function hideRowsPastDepth_1(){
    sheet.getRange(4, sheetGlobalsColumn).setValue(1);
    hideRowsPastDepth();
    //    reformatHierarchy();
};


function hideRowsPastDepth_2(){
    sheet.getRange(4, sheetGlobalsColumn).setValue(2);
    hideRowsPastDepth();
}

function hideRowsPastDepth_3(){
    sheet.getRange(4, sheetGlobalsColumn).setValue(3);
    hideRowsPastDepth();
}

function hideRowsPastDepth_4(){
    sheet.getRange(4, sheetGlobalsColumn).setValue(4);
    hideRowsPastDepth();
}

function hideRowsPastDepth_5(){
    sheet.getRange(4, sheetGlobalsColumn).setValue(5);
    hideRowsPastDepth();
}



function boldRow_atDepth(dep){
    sheet.getRange(5, sheetGlobalsColumn).setValue(dep);
    reformatHierarchy();
}

function boldRow_1(){ boldRow_atDepth(1) };
function boldRow_2(){ boldRow_atDepth(2) };
function boldRow_3(){ boldRow_atDepth(3) };
function boldRow_4(){ boldRow_atDepth(4) };
function boldRow_5(){ boldRow_atDepth(5) };
function boldRow_none(){ boldRow_atDepth('x') };
