package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.CsvOpenOptions;
import com.grapecity.documents.excel.OpenFileFormat;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ImportCsvFileToWorkbookWithOptions extends ExampleBase {

    @Override
    public void execute(Workbook workbook) { //Open csv with more settings.

        CsvOpenOptions options = new CsvOpenOptions();
        options.setColumnSeparator("-");

        //Change the path to the real file path when open.
        workbook.open("source.csv", options);
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
