package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class SaveWorkbookToPDF extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        //set style.
        sheet.getRange("A1").setValue("Sheet1");
        sheet.getRange("A1").getFont().setName("Wide Latin");
        sheet.getRange("A1").getFont().setColor(Color.getRed());
        sheet.getRange("A1").getInterior().setColor(Color.getGreen());

        //change the path to real export path when save.
        workbook.save("dest.pdf", SaveFileFormat.Pdf);
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