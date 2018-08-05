package com.grapecity.documents.excel.examples.features.tables;

import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.ITableStyle;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ApplyTableStyle extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //add table.
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        ITable table = worksheet.getTables().add(worksheet.getRange("A1:F7"), true);
        worksheet.getRange("A:F").setColumnWidth(15);

        //Add one custom table style.
        ITableStyle style = workbook.getTableStyles().add("test");
        //set custom table style for table.
        table.setTableStyle(style);

        //Use table style name get one build in table style.
        ITableStyle tableStyle = workbook.getTableStyles().get("TableStyleMedium3");
        //set built-in table style for table.
        table.setTableStyle(tableStyle);

    }

}
