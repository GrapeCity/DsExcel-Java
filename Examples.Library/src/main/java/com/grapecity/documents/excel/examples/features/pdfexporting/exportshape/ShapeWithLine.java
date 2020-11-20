package com.grapecity.documents.excel.examples.features.pdfexporting.exportshape;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.LineDashStyle;
import com.grapecity.documents.excel.drawing.LineStyle;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShapeWithLine extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        IShape rectangle = sheet.getShapes().addShape(AutoShapeType.Rectangle, 20, 50, 200, 200);
        rectangle.getLine().setDashStyle(LineDashStyle.Dash);
        rectangle.getLine().setStyle(LineStyle.Single);
        rectangle.getLine().setWeight(7);
        rectangle.getLine().getColor().setRGB(Color.GetYellow());

        IShape donut = sheet.getShapes().addShape(AutoShapeType.Donut, 260, 50, 200, 200);
        donut.getLine().setDashStyle(LineDashStyle.DashDotDot);
        donut.getLine().setStyle(LineStyle.Single);
        donut.getLine().setWeight(7);
        donut.getLine().getColor().setRGB(Color.GetRed());
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

}
