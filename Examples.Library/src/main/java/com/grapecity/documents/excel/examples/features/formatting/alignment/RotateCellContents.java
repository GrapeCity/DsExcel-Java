package com.grapecity.documents.excel.examples.features.formatting.alignment;

import com.grapecity.documents.excel.HorizontalAlignment;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.VerticalAlignment;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class RotateCellContents extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IRange rangeB2 = worksheet.getRange("B2");
        rangeB2.setValue("Rotated Cell Contents");
        rangeB2.setHorizontalAlignment(HorizontalAlignment.Center);
        rangeB2.setVerticalAlignment(VerticalAlignment.Center);
        // Rotate cell contents.
        rangeB2.setOrientation(15);

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
