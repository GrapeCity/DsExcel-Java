package com.grapecity.documents.excel.examples.features.formatting.fill;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SolidFill extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A2").getInterior().setColor(Color.GetGreen());
    }

}
