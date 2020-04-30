package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigureBestFitRowColumn extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        // Set text for some cells.
        sheet.getRange("A1").setValue("Grapecity");
        sheet.getRange("A2").setValue("Document For .NET");
        sheet.getRange("B1").setValue("Grapecity");
        sheet.getRange("B2").setValue("Excel for .NET");
        //Set font size of cell "A2"
        sheet.getRange("A2").getFont().setSize(20);

        // Set bestFitColumns/bestFitRows as true.
        sheet.getPageSetup().setBestFitColumns(true);
        sheet.getPageSetup().setBestFitRows(true);

        // Set print gridline and heading.
        sheet.getPageSetup().setPrintGridlines(true);
        sheet.getPageSetup().setPrintHeadings(true);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
	
}
