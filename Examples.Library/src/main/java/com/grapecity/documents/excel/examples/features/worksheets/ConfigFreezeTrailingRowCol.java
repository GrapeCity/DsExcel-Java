package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigFreezeTrailingRowCol extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        //Use sheet index to get worksheet.
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // freeze pane
        worksheet.freezePanes(2, 2);

        // freeze trailing pane
        worksheet.freezeTrailingPanes(2, 3);
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
