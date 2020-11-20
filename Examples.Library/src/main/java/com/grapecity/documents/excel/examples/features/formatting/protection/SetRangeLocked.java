package com.grapecity.documents.excel.examples.features.formatting.protection;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SetRangeLocked extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //config range B1's Locked property.
        worksheet.getRange("B1").setLocked(false);
        //protect worksheet, range B1 can be modified in exported xlsx file.
        worksheet.setProtection(true);
    }

    @Override
    public boolean getShowViewer() {

        return false;
    }

}
