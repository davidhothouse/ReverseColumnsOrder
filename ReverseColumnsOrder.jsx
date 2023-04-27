//DESCRIPTION: reverses the order of the affected columns.

// ** NOTE:
// Make sure there is no merged cells in affected columns.

// Created by Edmond Lo @ Hothouse Design
// hothousedesign.com.au

// Created 2023-04-19

// Make the script un-doable with one command
app.doScript(Main, undefined, undefined, UndoModes.ENTIRE_SCRIPT,"Run Script");

function Main() {
    // Check to see whether any InDesign documents are open.
    if (app.documents.length == 0) {
        alert("No InDesign documents are open. Please open a document and try again.");
        return;
    }

    // Check to see if exactly 1 item is selected
    if (app.selection.length != 1) {
        alert("Please select some cells/columns and try again.");
        return;
    }

	var myDoc = app.activeDocument;
	// Evaluate the selection based on its type.
	switch (app.selection[0].constructor.name){
		case "Table":
			reorderAllColumns(myDoc);
			break;
		case "Cell":
			reorderColumns(myDoc);
			break;
		default:
			alert("Please select some cells/columns and try again.");
			break;
	}
}


// Reverse the order of the columns of the selected cells
function reorderColumns(myDoc) {

	var fullRowNo = app.selection[0].parent.columns[0].cells.length; // Get the number of rows in the table
	var celSel = app.selection[0].columns.everyItem().getElements(); // Get the columns of the selected cells

	// Check for merged cells in affected columns
	// Iterate through each of the affected column
	for (var i=0; i<celSel.length; i++) {
		var selCol = celSel[i];
		if (selCol.cells.length != fullRowNo) { // The column has a different number of cells to the number of rows in the table
			alert("There are cells merged across columns in selection. Please unmerge and run the script again.");
			return;
		}

		// Iterate through each cell in the column
		for (var j=0; j<selCol.cells.length; j++) {
			var myCell = selCol.cells[j];
			if (myCell.columnSpan > 1) { // columnSpan is more than 1 if cells are merged
				alert("There are cells merged across columns in selection. Please unmerge and run the script again.");
				return;
			}
		}
	}

	var selColumns = myDoc.selection[0].columns;
	var selColNum = selColumns.length;
	var firstSelColNum = selColumns[0].index;
	var lastSelColNum = selColumns[-1].index;
	var newFirstSelCol, newLastSelCol;

	// Perform reversing task
	for (var k=selColNum-2; k>=0; k--) {
		newFirstSelCol = myDoc.selection[0].parent.columns[firstSelColNum + k];
		newLastSelCol = myDoc.selection[0].parent.columns[lastSelColNum + 1];
		selColumns.add(LocationOptions.AT_END);
		app.select(newFirstSelCol);
		app.copy();
		app.select(newLastSelCol);
		app.paste();
		newLastSelCol.width = newFirstSelCol.width;
		newFirstSelCol.remove();
	}
}

// Reverse the order of the columns of selected table
function reorderAllColumns(myDoc) {
	
	var celSel = app.selection[0].cells;

	// Check for merged cells
	// Iterate through each cell in the table
	for ( var c = 0; c < celSel.length; c++ ) { 
		if (celSel[c].columnSpan > 1) { // columnSpan is more than 1 if cells are merged
			alert("There are cells merged across columns in selection. Please unmerge and run the script again.");
			return;
		}
	}

	var selColumns = myDoc.selection[0].columns;
	var selColNum = selColumns.length;

	// Perform reversing task
    for (var i=selColNum-2; i>=0; i--) {
		selColumns.add(LocationOptions.AT_END);
		app.select(selColumns[i]);
		app.copy();
		app.select(selColumns[-1]);
		app.paste();
		selColumns[i].remove();
	}
}