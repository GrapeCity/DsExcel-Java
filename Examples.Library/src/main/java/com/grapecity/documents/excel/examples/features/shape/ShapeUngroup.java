package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.drawing.*;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShapeUngroup extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        IShapes shapes = worksheet.getShapes();
        IShape pentagon = shapes.addShape(AutoShapeType.RegularPentagon, 89.4, 57.0, 153.6, 90.6);
        IShape pie = shapes.addShape(AutoShapeType.Pie, 344.4, 156.8, 50.4, 60.0);
        IShapeRange shpRange = shapes.getRange(new String[] { pentagon.getName(), pie.getName() });

        // Group the shape range
        IShape grouped = shpRange.group();

        // Ungroup the group shape
        shpRange = grouped.ungroup();
    }

    @Override
    public boolean getShowViewer() {
        return true;
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
