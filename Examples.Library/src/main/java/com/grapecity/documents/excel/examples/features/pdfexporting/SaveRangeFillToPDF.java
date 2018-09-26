package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class SaveRangeFillToPDF extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("A1:C2").merge();
        worksheet.getRange("A1:C2").getInterior().setColor(Color.GetGreen());
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
}