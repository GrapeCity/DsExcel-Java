package com.grapecity.documents.excel.examples.features.formatting.fonts;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.UnderlineType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class FontUnderline extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("A1").setValue("Single Underline");
        worksheet.getRange("A1").getFont().setUnderline(UnderlineType.Single);
    }

}
