package com.grapecity.documents.excel.examples.features.formatting.fill;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.ILinearGradient;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Pattern;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class LinearGradientFill extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1").getInterior().setPattern(Pattern.LinearGradient);
        ((ILinearGradient) worksheet.getRange("A1").getInterior().getGradient()).getColorStops().get(0).setColor(Color.GetRed());
        ((ILinearGradient) worksheet.getRange("A1").getInterior().getGradient()).getColorStops().get(1).setColor(Color.GetYellow());

        ((ILinearGradient) worksheet.getRange("A1").getInterior().getGradient()).setDegree(90);

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
