package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.GradientStyle;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.PresetGradientType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigShapeWithGradientFill extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addShape(AutoShapeType.Parallelogram, 1, 1, 200, 100);
        shape.getFill().presetGradient(GradientStyle.Vertical, 3, PresetGradientType.Silver);
        shape.getFill().setRotateWithObject(false);

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
