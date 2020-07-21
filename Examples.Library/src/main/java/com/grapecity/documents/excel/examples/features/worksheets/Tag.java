package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class Tag extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Add tag for worksheet
        worksheet.setTag("This is a Tag for sheet.");

        //Add tag for range A1:B2
        worksheet.getRange("A1:B2").setTag("This is a Tag for A1:B2");

        //Add tag for row 4
        worksheet.getRange("A4").getEntireRow().setTag("This is a Tag for Row 4");

        //Add tag for column F
        worksheet.getRange("F5").getEntireColumn().setTag("This is a Tag for Column F");
    }

    @Override
    public boolean getCanDownload() {
        return false;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
}
