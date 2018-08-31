package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class ConfigurePageOrder extends ExampleBase {
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

        sheet.getPageSetup().setPrintHeadings(true);

        //Set page order. Now the page order is p1-p3-p2-p4. Origin is p1-p2-p3-p4.
        sheet.getPageSetup().setOrder(Order.OverThenDown);
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