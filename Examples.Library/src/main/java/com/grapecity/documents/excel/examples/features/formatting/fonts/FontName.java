package com.grapecity.documents.excel.examples.features.formatting.fonts;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ThemeFont;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FontName extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // use Name property to set font name.
        worksheet.getRange("A1").setValue("Calibri");
        worksheet.getRange("A1").getFont().setName("Calibri");

        // use ThemeFont property to set font name.
        worksheet.getRange("A2").setValue("Major theme font");
        worksheet.getRange("A2").getFont().setThemeFont(ThemeFont.Major);

    }

}
