package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class MergeCells extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //A1:C4 is a single merged cell.
        worksheet.getRange("A1:C4").merge();
        //H5:J5 is a single merged cell.
        //H6:J6 is a single merged cell.
        worksheet.getRange("H5:J6").merge(true);

        //select A1:H5's entire merge area A1:J5, entire merge area is a bounding rectangle.
        IRange entireMergeArea = worksheet.getRange("A1:H5").getEntireMergeArea();
        entireMergeArea.select();

        //judge if H5 is a merged cell.
        if (worksheet.getRange("J5").getMergeCells()) {
            //set value to the top left cell of the merge area.
            worksheet.getRange("J5").getEntireMergeArea().get(0, 0).setValue(1);
        }

    }

}
