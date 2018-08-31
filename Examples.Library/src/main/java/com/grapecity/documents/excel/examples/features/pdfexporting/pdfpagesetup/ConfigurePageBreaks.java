package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class ConfigurePageBreaks extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        sheet.getRange("A1:B5").setValue(new Object[][]
            {
                {1, 2},
                {3, 4},
                {5, 6},
                {7, 8},
                {9, 10}
            });

        //Add page break
        sheet.getHPageBreaks().add(sheet.getRange("B3"));
        sheet.getVPageBreaks().add(sheet.getRange("B3"));
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