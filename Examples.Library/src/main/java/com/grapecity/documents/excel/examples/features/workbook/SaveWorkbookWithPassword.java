package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.XlsxSaveOptions;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SaveWorkbookWithPassword extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Change the path to real export path when save.
        XlsxSaveOptions options = new XlsxSaveOptions();
        options.setPassword("123456");

        workbook.save("dest.xlsx", options);
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