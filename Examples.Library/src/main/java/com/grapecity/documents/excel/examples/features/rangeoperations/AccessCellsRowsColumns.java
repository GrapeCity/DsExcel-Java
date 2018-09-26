package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AccessCellsRowsColumns extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IRange range = worksheet.getRange("A5:B7");

        //set value for cell A7.
        range.getCells().get(4).setValue("A7");

        //cell is B6
        range.getCells().get(1, 1).setValue("B6");

        //row count is 3 and range is A6:B6.
        int rowCount = range.getRows().getCount();
        String row = range.getRows().get(1).toString();

        //set interior color for row range A6:B6.
        range.getRows().get(1).getInterior().setColor(Color.GetLightBlue());

        //column count is 2 and range is B5:B7.
        int columnCount = range.getColumns().getCount();
        String column = range.getColumns().get(1).toString();

        //set values for column range B5:B7.
        range.getColumns().get(1).getInterior().setColor(Color.GetLightSkyBlue());

        //entire rows are from row 5 to row 7
        String entirerow = range.getEntireRow().toString();

        //entire columns are from column A to column B
        String entireColumn = range.getEntireColumn().toString();

    }

}
