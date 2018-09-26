package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.XlsxOpenOptions;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ImportExcelFileWithoutCalculation extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //When XlsxOpenOptions.DoNotRecalculateAfterOpened means GrapeCity Documents for Excel will just read all the cached values without calculating again after
        //opening an Excel file.
        //Change the path to the real file path when open.
        XlsxOpenOptions options = new XlsxOpenOptions();
        options.setDoNotRecalculateAfterOpened(true);

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
