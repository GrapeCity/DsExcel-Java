package com.grapecity.documents.excel.examples.features.pdfexporting.text;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class NumberFormating extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        sheet.getRange("B3:B7").setValue(123456.789);
        sheet.getRange("B9:B13").setValue(-123456.789);

        sheet.getRange("B4, B10").setNumberFormat("0.00;[Red]0.00");
        sheet.getRange("B5, B11").setNumberFormat("$#,##0.00;[Red]$#,##0.00");
        sheet.getRange("B6, B12").setNumberFormat("0.00E+00");
        sheet.getRange("B7, B13").setNumberFormat("_($* #,##0.00_);_($* (#,##0.00);_($* \" - \"??_);_(@_)");

        //Other settings
        sheet.getColumns().get(1).setColumnWidthInPixel(100);
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