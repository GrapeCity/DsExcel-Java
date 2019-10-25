package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ControlAdjustingPageBreaks extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);
        sheet.getRange("A1:E5").setValue(new Object[][] { 
            { 1, 2, 3, 4, 5 }, 
            { 6, 7, 8, 9, 10 }, 
            { 11, 12, 13, 14, 15 },
            { 16, 17, 18, 19, 20 },
            { 21, 22, 23, 24, 25 }, 
        });

        // Add page break
        sheet.getHPageBreaks().add(sheet.getRange("D4")); // Add a horizontal page break before the fourth row.
        sheet.getVPageBreaks().add(sheet.getRange("D4")); // Add a vertical page break before the fourth column.

        // Delete rows and columns before the page breaks, the page breaks will be adjusted.
        sheet.getRange("1:1").delete(); // The hPageBreak is before the third row.
        sheet.getRange("A:A").delete(); // The vPageBreak is before the third column.

        // Set the page breaks are fixed, it will not be adjusted when inserting/deleting rows/columns.
        sheet.setFixedPageBreaks(true);

        sheet.getRange("1:1").delete(); // The hPageBreak is still before the third row.
        sheet.getRange("A:A").delete(); // The vPageBreak is still before the third column.

	}

	@Override
	public boolean getShowViewer() {
        return false;
	}

}
