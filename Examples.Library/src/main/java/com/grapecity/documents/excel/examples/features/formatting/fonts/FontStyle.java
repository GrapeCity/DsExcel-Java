package com.grapecity.documents.excel.examples.features.formatting.fonts;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FontStyle extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("A1").setValue("Bold");
        worksheet.getRange("A1").getFont().setBold(true);

        worksheet.getRange("A2").setValue("Italic");
        worksheet.getRange("A2").getFont().setItalic(true);

        worksheet.getRange("A3").setValue("Bold Italic");
        worksheet.getRange("A3").getFont().setBold(true);
        worksheet.getRange("A3").getFont().setItalic(true);
    }

}
