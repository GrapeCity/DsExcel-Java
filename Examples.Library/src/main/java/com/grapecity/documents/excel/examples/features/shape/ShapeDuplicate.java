package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShapeDuplicate extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        //Create shape
        IShape shape = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 50, 50, 200, 200);

        //Duplicate shape
        IShape newShape = shape.duplicate();
    }

    @Override
    public boolean getShowViewer() {

        return true;

    }

}
