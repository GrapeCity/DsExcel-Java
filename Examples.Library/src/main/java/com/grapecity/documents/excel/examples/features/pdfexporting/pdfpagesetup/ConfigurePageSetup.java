package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class ConfigurePageSetup extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        //Set data.
        sheet.getRange("A1:G10").setValue("Text");

        //Print rowheader and columnheader.
        sheet.getPageSetup().setPrintHeadings(true);

        //Print gridlines.
        sheet.getPageSetup().setPrintGridlines(true);
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