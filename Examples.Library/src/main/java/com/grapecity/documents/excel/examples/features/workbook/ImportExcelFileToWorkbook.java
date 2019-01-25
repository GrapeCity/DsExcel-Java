package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.OpenFileFormat;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ImportExcelFileToWorkbook extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Change the path to the real file path when open.
        workbook.open("source.xlsx");
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
