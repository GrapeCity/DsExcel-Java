package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IAdjustments;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigAdjustmentPoints extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addShape(AutoShapeType.RightArrowCallout, 20, 20, 200, 100);

        IAdjustments adjustments = shape.getAdjustments();
        adjustments.set(0, 0.5);// arrow neck width
        adjustments.set(1, 0.4);// arrow head width
        adjustments.set(2, 0.5);// arrow head height
        adjustments.set(3, 0.6);// text box width

    }

    @Override
    public boolean getIsNew() {

        return true;

    }
	
}
