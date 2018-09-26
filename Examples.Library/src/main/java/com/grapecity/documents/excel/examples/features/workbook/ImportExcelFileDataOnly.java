package com.grapecity.documents.excel.examples.features.workbook;

import java.util.EnumSet;

import com.grapecity.documents.excel.ImportFlags;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.XlsxOpenOptions;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ImportExcelFileDataOnly extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Use XlsxOpenOptions.ImportFlags to control what you want to import from excel, ImportFlags.Data means only the data will be imported
        //Change the path to the real file path when open.
        XlsxOpenOptions options = new XlsxOpenOptions();
        options.setImportFlags(EnumSet.of(ImportFlags.Data));

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
