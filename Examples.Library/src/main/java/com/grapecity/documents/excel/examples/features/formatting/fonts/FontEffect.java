package com.grapecity.documents.excel.examples.features.formatting.fonts;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FontEffect extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("A1").setValue("Strikethrough");
        worksheet.getRange("A1").getFont().setStrikethrough(true);

        worksheet.getRange("A2").setValue("Superscript");
        worksheet.getRange("A2").getFont().setSuperscript(true);

        worksheet.getRange("A3").setValue("Subscript");
        worksheet.getRange("A3").getFont().setSubscript(true);

    }

    @Override
    public boolean getShowViewer() {

        return false;
    }

    @Override
    public boolean getShowScreenshot() {

        return true;
    }

}
