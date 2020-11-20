package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class HiddenRowColumn extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("E1").setValue(1);

        //Hidden row 2:6.
        worksheet.getRange("2:6").setHidden(true);

        //Hidden column A:D.
        worksheet.getRange("A:D").setHidden(true);

    }

}
