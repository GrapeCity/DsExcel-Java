package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class GetWorksheetUsedRange extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("H6:M7").setValue(1);
        worksheet.getRange("J9:J10").merge();

        // set interior color for worksheet usedRange "H6:M10".
        IRange usedrange = worksheet.getUsedRange();
        usedrange.getInterior().setColor(Color.GetLightBlue());
    }

}
