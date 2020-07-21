package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigShowHVGridlinesSeparately extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        //Use sheet index to get worksheet.
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Set not to show horizontal gridlines
        worksheet.getSheetView().setDisplayHorizontalGridlines(false);

        // Set to show vertical gridlines
        worksheet.getSheetView().setDisplayVerticalGridlines(true);
    }

    @Override
    public boolean getCanDownload() {
        return false;
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
