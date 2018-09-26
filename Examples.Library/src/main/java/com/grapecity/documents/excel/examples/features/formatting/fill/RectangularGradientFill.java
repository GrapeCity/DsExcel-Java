package com.grapecity.documents.excel.examples.features.formatting.fill;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IRectangularGradient;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Pattern;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class RectangularGradientFill extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1").getInterior().setPattern(Pattern.RectangularGradient);
        ((IRectangularGradient) worksheet.getRange("A1").getInterior().getGradient()).getColorStops().get(0).setColor(Color.GetRed());
        ((IRectangularGradient) worksheet.getRange("A1").getInterior().getGradient()).getColorStops().get(1).setColor(Color.GetGreen());

        ((IRectangularGradient) worksheet.getRange("A1").getInterior().getGradient()).setBottom(0.2);
        ((IRectangularGradient) worksheet.getRange("A1").getInterior().getGradient()).setRight(0.3);
        ((IRectangularGradient) worksheet.getRange("A1").getInterior().getGradient()).setTop(0.4);
        ((IRectangularGradient) worksheet.getRange("A1").getInterior().getGradient()).setLeft(0.5);

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
