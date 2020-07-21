package com.grapecity.documents.excel.examples.features.pdfexporting.exportshape;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class BasicShape extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {
        
        IWorksheet sheet = workbook.getWorksheets().get(0);

        // Add a rectangle
        sheet.getShapes().addShape(AutoShapeType.Rectangle, 20, 10, 100, 100);

        // Add an oval
        sheet.getShapes().addShape(AutoShapeType.Oval, 20, 160, 100, 100);

        // Add a triangle
        sheet.getShapes().addShape(AutoShapeType.IsoscelesTriangle, 20, 310, 100, 100);

        // Add a "Not Allowed" symbol
        sheet.getShapes().addShape(AutoShapeType.NoSymbol, 20, 460, 100, 100);

        // Add a "Smile Face" symbol
        sheet.getShapes().addShape(AutoShapeType.SmileyFace, 20, 600, 100, 100);

        // Add a "Heart" symbol
        sheet.getShapes().addShape(AutoShapeType.Heart, 170, 10, 100, 100);

        // Add a "Sun" symbol
        sheet.getShapes().addShape(AutoShapeType.Sun, 170, 160, 100, 100);

        // Add a RightArrow
        sheet.getShapes().addShape(AutoShapeType.RightArrow, 170, 310, 100, 100);

        // Add a CurvedRightArrow
        sheet.getShapes().addShape(AutoShapeType.CurvedRightArrow, 170, 460, 100, 100);

        // Add a QuadArrow
        sheet.getShapes().addShape(AutoShapeType.QuadArrow, 170, 600, 100, 100);

        // Add a MathNotEqual
        sheet.getShapes().addShape(AutoShapeType.MathNotEqual, 320, 10, 100, 100);

        // Add a FlowchartMultidocument
        sheet.getShapes().addShape(AutoShapeType.FlowchartMultidocument, 320, 160, 100, 100);

        // Add a five points star
        sheet.getShapes().addShape(AutoShapeType.Shape5pointStar, 320, 310, 100, 100);

        // Add a CurvedUpRibbon
        sheet.getShapes().addShape(AutoShapeType.CurvedUpRibbon, 320, 460, 100, 100);

        // Add a OvalCallout
        sheet.getShapes().addShape(AutoShapeType.OvalCallout, 320, 580, 100, 100);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

}
