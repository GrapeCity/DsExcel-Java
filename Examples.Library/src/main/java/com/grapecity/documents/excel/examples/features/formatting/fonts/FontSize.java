package com.grapecity.documents.excel.examples.features.formatting.fonts;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FontSize extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("A1").setValue("Font size is 15");
        worksheet.getRange("A1").getFont().setSize(15);
    }

}
