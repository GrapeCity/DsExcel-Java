package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CutCopyRange extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IWorksheet worksheet2 = workbook.getWorksheets().add();

        worksheet.getRange("B3:D12").setValue(5);
        worksheet.getRange("B3:D12").getInterior().setColor(Color.GetLightGreen());

        //Copy
        worksheet.getRange("B3:D12").copy(worksheet.getRange("E5"));

        //Cut
        worksheet.getRange("B3:D12").cut(worksheet.getRange("I5:K14"));

        worksheet.getRange("I1:K2").setValue(2);
        worksheet.getRange("I1:K2").getInterior().setColor(Color.GetPink());

        //cross sheet cut copy.
        worksheet.getRange("I1:K2").cut(worksheet2.getRange("H5"));
        worksheet.getRange("G4:H5").copy(worksheet2.getRange("A1"));
    }

}
