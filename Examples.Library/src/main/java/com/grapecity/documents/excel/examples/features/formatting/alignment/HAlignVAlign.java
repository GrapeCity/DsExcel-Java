package com.grapecity.documents.excel.examples.features.formatting.alignment;

import com.grapecity.documents.excel.HorizontalAlignment;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.VerticalAlignment;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class HAlignVAlign extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getColumns().get(0).setColumnWidth(17);

        IRange rangeA1 = worksheet.getRange("A1");
        rangeA1.setValue("Right and top");
        rangeA1.setHorizontalAlignment(HorizontalAlignment.Right);
        rangeA1.setVerticalAlignment(VerticalAlignment.Top);

        IRange rangeA2 = worksheet.getRange("A2");
        rangeA2.setValue("Center");
        rangeA2.setHorizontalAlignment(HorizontalAlignment.Center);
        rangeA2.setVerticalAlignment(VerticalAlignment.Center);

        IRange rangeA3 = worksheet.getRange("A3");
        rangeA3.setValue("Left and bottom, indent");
        rangeA3.setIndentLevel(1);

        worksheet.getRange("A1:A3").setColumnWidth(50);
        worksheet.getRange("A1:A3").setRowHeight(30);
    }

}
