package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class ConfigurePrintArea extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        int row = 50;
        int column = 14;
        Object[][] data = new Object[row][column];
        for (int i = 0; i < row; i++) {
            for (int j = 0; j < column; j++) {
                data[i][j] = "R" + i + "C" + j;
            }
        }

        //Set data.
        sheet.getRange("A1:N50").setValue(data);
        sheet.getRange("C10:H20").setValue("PrintArea");
        sheet.getRange("C10:H20").getInterior().setColor(Color.GetYellow());
        sheet.getPageSetup().setPrintHeadings(true);

        //Set print area.
        sheet.getPageSetup().setPrintArea("$C$10:$H$20");
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