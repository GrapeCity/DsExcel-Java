package com.grapecity.documents.excel.examples.features.formatting.borders;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddingBordersToRange extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IRange rangeB2_E6 = worksheet.getRange("B2:E6");

        //set left, top, right, bottom borders together.
        rangeB2_E6.getBorders().setLineStyle(BorderLineStyle.DashDot);
        rangeB2_E6.getBorders().setColor(Color.GetGreen());

        //set inside horizontal border.
        rangeB2_E6.getBorders().get(BordersIndex.InsideHorizontal).setLineStyle(BorderLineStyle.Dashed);
        rangeB2_E6.getBorders().get(BordersIndex.InsideHorizontal).setColor(Color.GetTomato());

        //set inside vertical border.
        rangeB2_E6.getBorders().get(BordersIndex.InsideVertical).setLineStyle(BorderLineStyle.Double);
        rangeB2_E6.getBorders().get(BordersIndex.InsideVertical).setColor(Color.GetBlue());

        //set top border individually.
        rangeB2_E6.getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Medium);
        rangeB2_E6.getBorders().get(BordersIndex.EdgeTop).setColor(Color.GetRed());

    }

}
