package com.grapecity.documents.excel.examples.features.imageexporting;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

public class ExportShapeToImage extends ExampleBase {
    @Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/GroupShape.xlsx");
        workbook.open(fileStream);
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Export the shape to image
        worksheet.getShapes().get(0).toImage(outputStream, ImageType.PNG);
    }

    @Override
    public boolean getSaveAsImage() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/GroupShape.xlsx" };
    }
}
