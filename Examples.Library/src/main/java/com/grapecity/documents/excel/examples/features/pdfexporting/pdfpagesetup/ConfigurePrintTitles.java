package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class ConfigurePrintTitles extends ExampleBase {
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

        sheet.getRange("A2:N4").setValue("TitleRows");
        sheet.getRange("A2:N4").getInterior().setColor(Color.GetGreen());

        sheet.getRange("C1:E50").setValue("TitleColumns");
        sheet.getRange("C1:E50").getInterior().setColor(Color.GetYellow());

        sheet.getPageSetup().setPrintHeadings(true);

        //Set print titles.
        sheet.getPageSetup().setPrintTitleRows("$2:$4");
        sheet.getPageSetup().setPrintTitleColumns("$C:$E");
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