package com.grapecity.documents.excel.examples.features.pdfexporting.exportshape;

import java.io.IOException;
import java.io.InputStream;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.drawing.PresetTexture;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ShapeWithPictureOrTextureFill extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {
		
        IWorksheet sheet = workbook.getWorksheets().get(0);        
        
        try {
            // Get stream of picture
            InputStream stream1 = this.getResourceStream("logo.png");
        	
            // Add a rectangle
            IShape rectangle = sheet.getShapes().addShape(AutoShapeType.Rectangle, 20, 20, 250, 50);
            // Set picture fill
            rectangle.getFill().userPicture(stream1, ImageType.PNG);
            rectangle.getLine().setTransparency(1);

            // Get stream of picture
            InputStream stream2 = this.getResourceStream("logo.png");
	
            // Add a oval
            IShape oval = sheet.getShapes().addShape(AutoShapeType.Oval, 20, 90, 250, 50);
            // Set picture fill
            oval.getFill().userPicture(stream2, ImageType.PNG);
            oval.getLine().getColor().setRGB(Color.FromArgb(0x49129E));
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Add a five point star
        IShape star = sheet.getShapes().addShape(AutoShapeType.Shape5pointStar, 300, 20, 100, 100);
        // Set picture fill
        star.getFill().presetTextured(PresetTexture.WaterDroplets);
    }

    @Override
    public boolean getSavePdf() {
        return true;
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
