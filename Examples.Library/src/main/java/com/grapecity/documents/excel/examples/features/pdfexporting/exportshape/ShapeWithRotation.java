package com.grapecity.documents.excel.examples.features.pdfexporting.exportshape;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShapeWithRotation extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {
        
        IWorksheet sheet = workbook.getWorksheets().get(0);

        sheet.getRange("D2").setValue("rectangle with 30 degrees");
        sheet.getRange("H2").setValue("right arrow with 40 degrees");

        // Add a rectangle with rotation
        IShape rectangle = sheet.getShapes().addShape(AutoShapeType.Rectangle, 50, 50, 200, 50);
        rectangle.setRotation(30);

        // Add a right arrow with rotation
        IShape rightArrowWithRotation = sheet.getShapes().addShape(AutoShapeType.RightArrow, 270, 50, 200, 100);
        rightArrowWithRotation.setRotation(40);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

}
