package com.grapecity.documents.excel.examples.features.formatting.alignment;

import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ReadingOrder;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class TextOrientation extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IRange rangeC1 = worksheet.getRange("C1");
        rangeC1.setValue("The ReadingOrder property is applied to set text direction.");
        rangeC1.setReadingOrder(ReadingOrder.RightToLeft);

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
