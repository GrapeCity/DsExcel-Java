package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.chart.AutoShapeType;
import com.grapecity.documents.excel.drawing.chart.LineDashStyle;
import com.grapecity.documents.excel.drawing.chart.LineStyle;
import com.grapecity.documents.excel.examples.ExampleBase;
import com.grapecity.documents.excel.style.ThemeColor;

public class ConfigShapeLine extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addShape(AutoShapeType.Parallelogram, 1, 1, 200, 100);
        shape.getLine().setDashStyle(LineDashStyle.Dash);
        shape.getLine().setStyle(LineStyle.Single);
        shape.getLine().setWeight(2);
        shape.getLine().getColor().setObjectThemeColor(ThemeColor.Accent6);
        shape.getLine().setTransparency(0.3);

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
