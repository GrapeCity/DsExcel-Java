package com.grapecity.documents.excel.examples.features.formatting.fill;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Pattern;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class PatternFill extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getRange("A1").getInterior().setPattern(Pattern.LightDown);
        worksheet.getRange("A1").getInterior().setColor(Color.GetPink());
        worksheet.getRange("A1").getInterior().setPatternColorIndex(5);

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
