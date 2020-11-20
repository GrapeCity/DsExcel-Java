package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class InsertDeleteRowsColumns extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet1 = workbook.getWorksheets().get(0);
        IWorksheet worksheet2 = workbook.getWorksheets().add();

        Object data = new Object[][]
                {
                        {1, 2, 3},
                        {4, 5, 6},
                        {7, 8, 9}
                };

        worksheet1.getRange("A1:C3").setValue(data);
        worksheet2.getRange("A1:C3").setValue(data);

        //Insert rows
        worksheet1.getRange("A2:B2").getEntireRow().insert();
        worksheet1.getRange("3:5").insert();

        //Insert columns
        worksheet1.getRange("B3:B5").getEntireColumn().insert();
        worksheet1.getRange("A:A").insert();

        //Delete rows
        worksheet2.getRange("A3:A5").getEntireRow().delete();
        worksheet2.getRange("2:4").delete();

        //Delete columns
        worksheet2.getRange("B3:B5").getEntireColumn().delete();
        worksheet2.getRange("A:A").delete();

    }

}
