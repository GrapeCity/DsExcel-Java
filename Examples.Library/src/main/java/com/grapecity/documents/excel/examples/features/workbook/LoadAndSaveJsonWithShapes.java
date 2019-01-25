package com.grapecity.documents.excel.examples.features.workbook;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.ThemeColor;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.XlsxOpenOptions;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.LineDashStyle;
import com.grapecity.documents.excel.drawing.LineStyle;
import com.grapecity.documents.excel.examples.ExampleBase;

public class LoadAndSaveJsonWithShapes extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        Workbook workbookWithShape = new Workbook();
        IWorksheet worksheet = workbookWithShape.getWorksheets().get(0);

        // Add a shape in worksheet
        IShape shape = worksheet.getShapes().addShape(AutoShapeType.Parallelogram, 1, 1, 200, 100);
        shape.getLine().setDashStyle(LineDashStyle.Dash);
        shape.getLine().setStyle(LineStyle.Single);
        shape.getLine().setWeight(2);
        shape.getLine().getColor().setObjectThemeColor(ThemeColor.Accent6);
        shape.getLine().setTransparency(0.3);

        // jsongString contains shapes
        String jsonString = workbookWithShape.toJson();

        // GcExcel can load json strig contains shapes now
        workbook.fromJson(jsonString);
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}
