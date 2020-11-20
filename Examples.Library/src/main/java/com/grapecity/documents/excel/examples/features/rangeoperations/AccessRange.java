package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AccessRange extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //use index to access cell A1.
        worksheet.getRange(0, 0).getInterior().setColor(Color.GetLightGreen());

        //use index to access range A1:B2
        worksheet.getRange(0, 0, 2, 2).setValue(5);

        //use string to access range.
        worksheet.getRange("A2").getInterior().setColor(Color.GetLightYellow());
        worksheet.getRange("C3:D4").getInterior().setColor(Color.GetTomato());
        worksheet.getRange("A5:B7, C3, H5:N6").setValue(2);

        //use index to access rows
        worksheet.getRows().get(2).getInterior().setColor(Color.GetLightSalmon());

        //use string to access rows
        worksheet.getRange("4:4").getInterior().setColor(Color.GetLightSkyBlue());

        //use index to access columns
        worksheet.getColumns().get(2).getInterior().setColor(Color.GetLightSalmon());

        //use string to access columns
        worksheet.getRange("D:D").getInterior().setColor(Color.GetLightSkyBlue());

        //use Cells to access range.
        worksheet.getCells().get(5).getInterior().setColor(Color.GetLightBlue());
        worksheet.getCells().get(5, 5).getInterior().setColor(Color.GetLightYellow());

        //access all rows in worksheet
        String allRows = worksheet.getRows().toString();

        //access all columns in worksheet
        String allColumns = worksheet.getColumns().toString();

        //access the entire sheet range
        String entireSheet = worksheet.getCells().toString();

    }

}
