package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.DeleteShiftDirection;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.InsertShiftDirection;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class InsertCellsDeleteCells extends ExampleBase {

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

        //Insert cells
        worksheet1.getRange("A2").insert();//Auto
        worksheet1.getRange("B2").insert(InsertShiftDirection.Down);
        worksheet1.getRange("C2").insert(InsertShiftDirection.Right);

        //Delete cells
        worksheet2.getRange("A2").delete();//Auto
        worksheet2.getRange("B2").delete(DeleteShiftDirection.Left);
        worksheet2.getRange("C2").delete(DeleteShiftDirection.Up);

    }

}
