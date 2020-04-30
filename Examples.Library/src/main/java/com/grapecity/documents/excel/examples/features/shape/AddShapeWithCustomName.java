package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class AddShapeWithCustomName extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addShape("custom parallelogram", AutoShapeType.Parallelogram, 1, 1, 200, 100);

        // Get shape by name
        IShape parallelogram = worksheet.getShapes().get("custom parallelogram");
        parallelogram.getFill().getColor().setRGB(Color.GetRed());
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
