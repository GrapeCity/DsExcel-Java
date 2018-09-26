package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Visibility;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigWorksheet extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Set worksheet tab color.
        worksheet.setTabColor(Color.GetGreen());

        //Set worksheet default row height.
        worksheet.setStandardHeight(20);
        //Set worksheet default column width.
        worksheet.setStandardWidth(50);

        //Split worksheet to panes.
        worksheet.splitPanes(worksheet.getRange("B3").getRow(), worksheet.getRange("B3").getColumn());

        IWorksheet worksheet1 = workbook.getWorksheets().add();
        //Hide worksheet.
        worksheet1.setVisible(Visibility.Hidden);

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public boolean getShowScreenshot() {

        return true;

    }

}
