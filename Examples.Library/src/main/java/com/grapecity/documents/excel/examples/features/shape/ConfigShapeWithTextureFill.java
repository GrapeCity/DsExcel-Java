package com.grapecity.documents.excel.examples.features.shape;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.PresetTexture;
import com.grapecity.documents.excel.drawing.TextureAlignment;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigShapeWithTextureFill extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addShape(AutoShapeType.Parallelogram, 1, 1, 200, 100);
        shape.getFill().presetTextured(PresetTexture.Canvas);
        shape.getFill().setTextureAlignment(TextureAlignment.Center);
        shape.getFill().setTextureOffsetX(2.5);
        shape.getFill().setTextureOffsetY(3.2);
        shape.getFill().setTextureHorizontalScale(0.9);
        shape.getFill().setTextureVerticalScale(0.2);
        shape.getFill().setTransparency(0.5);

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public boolean getShowScreenshot() {

        return true;

    }

}
