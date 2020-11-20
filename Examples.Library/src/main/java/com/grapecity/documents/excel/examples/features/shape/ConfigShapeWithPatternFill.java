package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.PatternType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigShapeWithPatternFill extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addShape(AutoShapeType.Parallelogram, 1, 1, 200, 100);
        shape.getFill().patterned(PatternType.Percent10);
        shape.getFill().getColor().setObjectThemeColor(ThemeColor.Accent2);
        shape.getFill().getPatternColor().setObjectThemeColor(ThemeColor.Accent6);

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
