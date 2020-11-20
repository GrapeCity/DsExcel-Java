package com.grapecity.documents.excel.examples.features.pdfexporting.exportshape;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SetShapeNotToPrint extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        // Add a rectangle
        IShape rectangle = sheet.getShapes().addShape(AutoShapeType.Rectangle, 20, 15, 100, 100);
        
        // Add an oval
        IShape oval = sheet.getShapes().addShape(AutoShapeType.Oval, 160, 15, 100, 100);

        //set oval not to print
        oval.setIsPrintable(false);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

}
