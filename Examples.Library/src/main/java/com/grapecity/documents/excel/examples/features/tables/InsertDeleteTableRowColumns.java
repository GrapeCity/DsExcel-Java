package com.grapecity.documents.excel.examples.features.tables;

import java.util.GregorianCalendar;

import com.grapecity.documents.excel.ITable;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class InsertDeleteTableRowColumns extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        Object data = new Object[][]{
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", new GregorianCalendar(1968, 5, 8), "Blue", 67, 165},
                {"Nia", "New York", new GregorianCalendar(1972, 6, 3), "Brown", 62, 134},
                {"Jared", "New York", new GregorianCalendar(1964, 2, 2), "Hazel", 72, 180},
                {"Natalie", "Washington", new GregorianCalendar(1972, 7, 8), "Blue", 66, 163},
                {"Damon", "Washington", new GregorianCalendar(1986, 1, 2), "Hazel", 76, 176},
                {"Angela", "Washington", new GregorianCalendar(1993, 1, 15), "Brown", 68, 145}
        };

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1:F7").setValue(data);
        worksheet.getRange("A:F").setColumnWidth(15);

        // add table.
        ITable table = worksheet.getTables().add(worksheet.getRange("A1:F7"), true);

        // add table column before first column.
        table.getColumns().add(0);
        // add table column before second column.
        table.getColumns().add(1);

        // delete first table column.
        table.getColumns().get(0).delete();
        // delete "City" table column.
        table.getColumns().get("City").delete();

        // insert a table row in table's last row.
        table.getRows().add();
        // delete second table row.
        table.getRows().get(1).delete();

    }

}
