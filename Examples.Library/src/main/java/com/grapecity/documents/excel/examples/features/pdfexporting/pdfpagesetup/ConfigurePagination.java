package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class ConfigurePagination extends ExampleBase {
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

        //Set data
        sheet.getRange("A1:N50").setValue(data);

        //Set paper size
        sheet.getPageSetup().setPaperSize(PaperSize.A5);

        //With API the margin's unit is pound, with Excel the margins display unit is inch.
        //One inch eaquals to 72 pounds. If the top margin is 0.8 inch, then please set PageSetup.TopMargin = 0.8*72(57.6);
        sheet.getPageSetup().setTopMargin(57.6);

        //Top margin in excel is 0.8 inch
        sheet.getPageSetup().setBottomMargin(21.6);
        sheet.getPageSetup().setLeftMargin(28.8);
        sheet.getPageSetup().setRightMargin(0);
        sheet.getPageSetup().setHeaderMargin(0);
        sheet.getPageSetup().setFooterMargin(93.6);
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