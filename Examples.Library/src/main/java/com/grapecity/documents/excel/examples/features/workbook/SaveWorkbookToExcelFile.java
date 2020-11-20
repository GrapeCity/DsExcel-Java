package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.SaveFileFormat;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SaveWorkbookToExcelFile extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //change the path to real export path when save.
        workbook.save("dest.xlsx");
    }

    @Override
    public boolean getCanDownload() {
        return false;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
}
