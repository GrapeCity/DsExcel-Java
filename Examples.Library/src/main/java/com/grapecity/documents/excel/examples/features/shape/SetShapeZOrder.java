package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.IShapes;
import com.grapecity.documents.excel.drawing.ZOrderType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SetShapeZOrder extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShapes shapes = worksheet.getShapes();

        IShape rectangle = shapes.addShape(AutoShapeType.Rectangle, 20, 20, 100, 100);
        rectangle.getFill().getColor().setRGB(Color.FromArgb(169, 209, 142));

        IShape oval = shapes.addShape(AutoShapeType.Oval, 50, 50, 100, 100);
        oval.getFill().getColor().setRGB(Color.FromArgb(157, 195, 230));

        IShape triangle = shapes.addShape(AutoShapeType.IsoscelesTriangle, 80, 80, 100, 100);
        triangle.getFill().getColor().setRGB(Color.FromArgb(255, 230, 153));

        // Set rectangle above oval
        rectangle.zOrder(ZOrderType.BringForward);

        // Set triangle to bottom
        triangle.zOrder(ZOrderType.SendToBack);
    }
}
