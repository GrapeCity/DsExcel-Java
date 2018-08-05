package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.style.color.Color;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class CutCopyRange extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        //TODO cut/copy Example
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IWorksheet worksheet2 = workbook.getWorksheets().add();

        worksheet.getRange("B3:D12").setValue(5);
        worksheet.getRange("B3:D12").getInterior().setColor(Color.getLightGreen());
    }

}
