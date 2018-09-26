package com.grapecity.documents.excel.examples.features.formatting.borders;

import com.grapecity.documents.excel.BorderLineStyle;
import com.grapecity.documents.excel.BordersIndex;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddBordersToCell extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IRange rangeB2 = worksheet.getRange("B2");

        //set left, top, right, bottom borders together.
        rangeB2.getBorders().setLineStyle(BorderLineStyle.DashDot);
        rangeB2.getBorders().setColor(Color.GetGreen());

        //set top border individually.
        rangeB2.getBorders().get(BordersIndex.EdgeTop).setLineStyle(BorderLineStyle.Medium);
        rangeB2.getBorders().get(BordersIndex.EdgeTop).setColor(Color.GetRed());

        //set diagonal down border individually.
        rangeB2.getBorders().get(BordersIndex.DiagonalDown).setLineStyle(BorderLineStyle.Hair);
        rangeB2.getBorders().get(BordersIndex.DiagonalDown).setColor(Color.GetBlue());

        //set diagonal up border individually.
        rangeB2.getBorders().get(BordersIndex.DiagonalUp).setLineStyle(BorderLineStyle.Dotted);
        rangeB2.getBorders().get(BordersIndex.DiagonalUp).setColor(Color.GetBlue());

    }

    @Override
    public boolean getShowViewer() {

        return false;
    }

    @Override
    public boolean getShowScreenshot() {

        return true;
    }
}
