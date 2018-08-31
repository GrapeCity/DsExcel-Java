package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.XlsxOpenOptions;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ImportExcelFileWithPassword extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        //Change the path to the real file path when open.
        XlsxOpenOptions options = new XlsxOpenOptions();
        options.setPassword("123456");

        workbook.open("source.xlsx", options);
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
