package com.grapecity.documents.excel.examples.features.worksheets;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AccessWorksheet extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        //Use sheet index to get worksheet.
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Use sheet name to get worksheet.
        IWorksheet worksheet1 = workbook.getWorksheets().get("Sheet1");

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

}
