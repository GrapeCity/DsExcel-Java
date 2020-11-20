package com.grapecity.documents.excel.examples.features.rangeoperations;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AccessEachAreaOfRange extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // set interior color for area1 A5:B7.
        IRange area1 = worksheet.getRange("A5:B7, C3, H5:N6").getAreas().getArea(0);
        area1.getInterior().setColor(Color.GetPink());

        // set interior color for area2 C3.
        IRange area2 = worksheet.getRange("A5:B7, C3, H5:N6").getAreas().getArea(1);
        area2.getInterior().setColor(Color.GetLightGreen());

        // set interior color for area3 H5:N6.
        IRange area3 = worksheet.getRange("A5:B7, C3, H5:N6").getAreas().getArea(2);
        area3.getInterior().setColor(Color.GetLightBlue());

    }

}
