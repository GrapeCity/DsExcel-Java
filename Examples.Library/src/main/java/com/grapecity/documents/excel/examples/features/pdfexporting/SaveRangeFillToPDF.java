package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

public class SaveRangeFillToPDF extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IRange rangeA1B2 = worksheet.getRange("A1:B2");
        rangeA1B2.merge();
        rangeA1B2.getInterior().setPattern(Pattern.LinearGradient);
        Object tempVar = rangeA1B2.getInterior().getGradient();
        ((ILinearGradient) ((tempVar instanceof ILinearGradient) ? tempVar : null)).getColorStops().get(0).setColor(Color.getRed());
        Object tempVar2 = rangeA1B2.getInterior().getGradient();
        ((ILinearGradient) ((tempVar2 instanceof ILinearGradient) ? tempVar2 : null)).getColorStops().get(1).setColor(Color.getYellow());
        Object tempVar3 = rangeA1B2.getInterior().getGradient();
        ((ILinearGradient) ((tempVar3 instanceof ILinearGradient) ? tempVar3 : null)).setDegree(90);

        IRange rangeE1E2 = worksheet.getRange("D1:E2");
        rangeE1E2.merge();
        rangeE1E2.getInterior().setPattern(Pattern.LightDown);
        rangeE1E2.getInterior().setColor(Color.getPink());
        rangeE1E2.getInterior().setPatternColorIndex(5);

        IRange rangeG1H2 = worksheet.getRange("G1:H2");
        rangeG1H2.merge();
        rangeG1H2.getInterior().setPattern(Pattern.RectangularGradient);
        Object tempVar4 = rangeG1H2.getInterior().getGradient();
        ((IRectangularGradient) ((tempVar4 instanceof IRectangularGradient) ? tempVar4 : null)).getColorStops().get(0).setColor(Color.getRed());
        Object tempVar5 = rangeG1H2.getInterior().getGradient();
        ((IRectangularGradient) ((tempVar5 instanceof IRectangularGradient) ? tempVar5 : null)).getColorStops().get(1).setColor(Color.getGreen());

        Object tempVar6 = rangeG1H2.getInterior().getGradient();
        ((IRectangularGradient) ((tempVar6 instanceof IRectangularGradient) ? tempVar6 : null)).setBottom(0.2);
        Object tempVar7 = rangeG1H2.getInterior().getGradient();
        ((IRectangularGradient) ((tempVar7 instanceof IRectangularGradient) ? tempVar7 : null)).setRight(0.3);
        Object tempVar8 = rangeG1H2.getInterior().getGradient();
        ((IRectangularGradient) ((tempVar8 instanceof IRectangularGradient) ? tempVar8 : null)).setTop(0.4);
        Object tempVar9 = rangeG1H2.getInterior().getGradient();
        ((IRectangularGradient) ((tempVar9 instanceof IRectangularGradient) ? tempVar9 : null)).setLeft(0.5);

        worksheet.getRange("J1:K2").merge();
        worksheet.getRange("J1:K2").getInterior().setColor(Color.getGreen());
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
}