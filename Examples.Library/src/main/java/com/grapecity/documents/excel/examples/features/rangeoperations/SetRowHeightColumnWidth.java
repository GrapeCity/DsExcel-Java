package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SetRowHeightColumnWidth extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //set row height for row 1:2.
        worksheet.getRange("1:2").setRowHeight(50);

        //set column width for column C:D.
        worksheet.getRange("C:D").setColumnWidth(20);
    }

}
