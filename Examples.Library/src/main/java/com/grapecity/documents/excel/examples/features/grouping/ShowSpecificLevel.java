package com.grapecity.documents.excel.examples.features.grouping;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShowSpecificLevel extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A:N").group();
        worksheet.getRange("A:F").group();
        worksheet.getRange("A:C").group();

        worksheet.getRange("Q:Z").group();
        worksheet.getRange("Q:T").group();

        //level 3 and level 4 will be collapsed. level 2 and level 1 expand.
        worksheet.getOutline().showLevels(0, 2);

    }

}
