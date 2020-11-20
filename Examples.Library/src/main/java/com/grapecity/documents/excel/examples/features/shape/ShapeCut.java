package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShapeCut extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        //Create a shape in worksheet, shape's range is Range["A7:B7"]
        IShape shape = worksheet.getShapes().addShape(AutoShapeType.Rectangle, 1, 1, 100, 100);
        //Range["A1:D10"] must contain Range["A7:B7"], cut a new shape to Range["C1:F7"]
        worksheet.getRange("A1:D10").cut(worksheet.getRange("E1"));
        worksheet.getRange("A1:D10").cut(worksheet.getRange("E1:I9"));

        //Cross sheet cut, cut a new shape to worksheet2's Range["C1:F7"]
        //IWorksheet worksheet2 = workbook.getWorksheets().add();
        //worksheet.getRange("A1:D10").cut(worksheet2.getRange("E1"));
        //worksheet.getRange("A1:D10").cut(worksheet2.getRange("E1:I9"));
    }

    @Override
    public boolean getShowViewer() {

        return true;

    }

}
