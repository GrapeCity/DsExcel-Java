package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class ConfigurePrintPageRange extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        //Set pages' data.
        sheet.getRange("A1:J46").setValue("Page1");
        sheet.getRange("A1:J46").getInterior().setColor(Color.getLightGreen());

        sheet.getRange("A47:J92").setValue("Page2");
        sheet.getRange("A47:J92").getInterior().setColor(Color.getLightYellow());

        sheet.getRange("K1:T46").setValue("Page3");
        sheet.getRange("K1:T46").getInterior().setColor(Color.getOrangeRed());

        sheet.getRange("K47:T92").setValue("Page4");
        sheet.getRange("K47:T92").getInterior().setColor(Color.getDarkOrange());

        sheet.getRange("U1:AD46").setValue("Page5");
        sheet.getRange("U1:AD46").getInterior().setColor(Color.getLightBlue());

        sheet.getRange("U47:AD92").setValue("Page6");
        sheet.getRange("U47:AD92").getInterior().setColor(Color.getIndianRed());
        sheet.getPageSetup().setPrintHeadings(true);

        //Set print page range, print p1, p3 to p5.
        sheet.getPageSetup().setPrintPageRange("1,3-5");
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