package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CellInfo extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // cell's value B2
        String cell = com.grapecity.documents.excel.CellInfo.CellIndexToName(1, 1);
        worksheet.getRange(cell).getInterior().setColor(Color.GetLightBlue());

        // rowIndex is 3 and columnIndex is 2
        int[] index = com.grapecity.documents.excel.CellInfo.CellNameToIndex("C4");
        int rowIndex = index[0];
        int columnIndex = index[1];
        worksheet.getRange(rowIndex, columnIndex).getInterior().setColor(Color.GetLightCoral());

        // column is D
        String column = com.grapecity.documents.excel.CellInfo.ColumnIndexToName(3);
        worksheet.getRange(String.format("%s:%s", column, column)).getInterior().setColor(Color.GetLightGreen());

        // columnIndex is 4
        columnIndex = com.grapecity.documents.excel.CellInfo.ColumnNameToIndex("E");
        worksheet.getColumns().get(columnIndex).getInterior().setColor(Color.GetLightSalmon());

        // row is 3
        String row = com.grapecity.documents.excel.CellInfo.RowIndexToName(2);
        worksheet.getRange(String.format("%s:%s", row, row)).getInterior().setColor(Color.GetLightSteelBlue());

        // rowIndex is 4
        rowIndex = com.grapecity.documents.excel.CellInfo.RowNameToIndex("5");
        worksheet.getRows().get(rowIndex).getInterior().setColor(Color.GetLightSkyBlue());
    }
}
