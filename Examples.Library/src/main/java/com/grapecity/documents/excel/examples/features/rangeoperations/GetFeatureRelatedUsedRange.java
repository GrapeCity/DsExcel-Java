package com.grapecity.documents.excel.examples.features.rangeoperations;

import java.util.EnumSet;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.UsedRangeType;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class GetFeatureRelatedUsedRange extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getRange("A1:B2").setValue(new Object[][]{
                {1, 2},
                {"aaa", "bbb"}
        });
        worksheet.getRange("A2:C3").getInterior().setColor(Color.GetGreen());

        //style used range is A2:C3.
        IRange usedRange_style = worksheet.getUsedRange(EnumSet.of(UsedRangeType.Style));
        usedRange_style.getInterior().setColor(Color.GetLightBlue());

    }

}
