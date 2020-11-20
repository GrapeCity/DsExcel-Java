package com.grapecity.documents.excel.examples.features.imageexporting;

import java.io.ByteArrayOutputStream;

import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.AutoShapeType;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ConvertShapeToImage extends ExampleBase {
	@Override
	public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Add a oval
        IShape oval = worksheet.getShapes().addShape(AutoShapeType.Oval, 0, 0, 191, 194);

        // Set soild fill for oval
        oval.getFill().getColor().setRGB(Color.GetOrangeRed());

        // Save the shape as image to a stream.
        oval.toImage(outputStream, ImageType.PNG);
	}

	@Override
	public boolean getSaveAsImage() {
        return true;
	}
}
