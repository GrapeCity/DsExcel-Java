package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigFrozenLineColor extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        //Use sheet index to get worksheet.
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // freeze pane
        worksheet.freezePanes(5, 5);

        // Set frozen line color as dark blue.
        worksheet.setFrozenLineColor(Color.GetDarkBlue());
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
