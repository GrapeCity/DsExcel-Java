package com.grapecity.documents.excel.examples.features.formatting.alignment;

import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShrinkToFit extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IRange rangeB1 = worksheet.getRange("B1");
        rangeB1.setValue("The ShrinkToFit property is applied");
        rangeB1.setShrinkToFit(true);

    }

}
