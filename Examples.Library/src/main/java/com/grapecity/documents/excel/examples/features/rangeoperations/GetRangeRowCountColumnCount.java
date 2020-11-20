package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class GetRangeRowCountColumnCount extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IRange range = worksheet.getRange("A5:B7");

        // cell count is 6.
        int cellcount = range.getCount();
        // cell count is 6.
        int cellcount1 = range.getCells().getCount();
        // row count is 3.
        int rowcount = range.getRows().getCount();
        // column count is 2.
        int columncount = range.getColumns().getCount();

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
