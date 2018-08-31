package com.grapecity.documents.excel.examples.features.pdfexporting.text;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class Overflow extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        sheet.getRange("F2, F4").setValue("This is a test string of overflow");

        sheet.getRange("F6, F8").setValue("This is a test string of overflow with right alignment");
        sheet.getRange("F6, F8").setHorizontalAlignment(HorizontalAlignment.Right);

        sheet.getRange("D8, H4").setValue(123);

        //Other settings
        sheet.getRange("A1:J10").getBorders().setLineStyle(BorderLineStyle.Thin);
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