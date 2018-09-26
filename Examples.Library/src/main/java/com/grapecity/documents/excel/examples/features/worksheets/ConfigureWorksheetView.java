package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.IWorksheetView;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigureWorksheetView extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Worksheet view settings.
        IWorksheetView sheetView = worksheet.getSheetView();
        sheetView.setDisplayFormulas(false);
        sheetView.setDisplayRightToLeft(true);
        sheetView.setGridlineColor(Color.GetRed());
        sheetView.setZoom(200);

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
