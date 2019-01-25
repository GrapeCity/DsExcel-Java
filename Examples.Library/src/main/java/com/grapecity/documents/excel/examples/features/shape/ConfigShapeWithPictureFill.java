package com.grapecity.documents.excel.examples.features.shape;

import java.io.IOException;
import java.io.InputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.drawing.PictureColorType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConfigShapeWithPictureFill extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);
        IShape shape = worksheet.getShapes().addShape(AutoShapeType.Parallelogram, 20, 20, 200, 100);
        InputStream stream = this.getResourceStream("logo.png");

        try {
            shape.getFill().userPicture(stream, ImageType.JPG);
        } catch (IOException e) {
            e.printStackTrace();
        }
        shape.getPictureFormat().setColorType(PictureColorType.Grayscale);
        shape.getPictureFormat().setBrightness(0.6);
        shape.getPictureFormat().setContrast(0.3);
        shape.getPictureFormat().getCrop().setPictureOffsetX(10);
        shape.getPictureFormat().getCrop().setPictureOffsetY(-5);
        shape.getPictureFormat().getCrop().setPictureWidth(120);
        shape.getPictureFormat().getCrop().setPictureHeight(80);

    }

    @Override
    public boolean getShowViewer() {

        return false;

    }

    @Override
    public boolean getShowScreenshot() {

        return true;

    }

    @Override
    public String[] getResources() {
        return new String[]{"logo.png"};
    }
}
