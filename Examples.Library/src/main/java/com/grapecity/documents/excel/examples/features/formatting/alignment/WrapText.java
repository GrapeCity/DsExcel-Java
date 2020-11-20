package com.grapecity.documents.excel.examples.features.formatting.alignment;

import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class WrapText extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IRange rangeB3 = worksheet.getRange("B3");
        rangeB3.setValue("The WrapText property is applied to wrap the text within a cell");
        rangeB3.setWrapText(true);

        worksheet.getRows().get(2).setRowHeight(150);
    }

}
