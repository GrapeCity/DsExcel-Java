package com.grapecity.documents.excel.examples.features.formatting.numberformat;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CustomNumberFormat extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        // Set range values.
        worksheet.getRange("A2:B2").setValue(-15.50);
        worksheet.getRange("A3:B3").setValue(555);
        worksheet.getRange("A4:B4").setValue(0);
        worksheet.getRange("A5:B5").setValue("Name");

        // Apply custom number format.
        worksheet.getRange("B2:B5").setNumberFormat("[Green]#.00;[Red]#.00;[Blue]0.00;[Cyan]\"product: \"@");
    }

}
