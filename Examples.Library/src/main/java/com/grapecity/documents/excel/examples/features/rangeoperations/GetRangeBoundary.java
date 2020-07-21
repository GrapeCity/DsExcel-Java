package com.grapecity.documents.excel.examples.features.rangeoperations;

import java.io.IOException;
import java.io.InputStream;

import com.grapecity.documents.excel.IRange;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Rectangle;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class GetRangeBoundary extends ExampleBase {

	@Override
    public void execute(Workbook workbook) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/Sport sign-up sheet.xlsx");
        workbook.open(fileStream);
        
        IWorksheet worksheet = workbook.getWorksheets().get(0);
        
        // Get the absolute location and size of the Range["G1"] in the worksheet.
        IRange range = worksheet.getRange("G1");
        Rectangle rect = com.grapecity.documents.excel.CellInfo.GetRangeBoundary(range);
        // Add the image to the Range["G1"]
        InputStream stream = this.getResourceStream("logo.png");
        try {
            worksheet.getShapes().addPictureInPixel(stream, ImageType.PNG, rect.getX(), rect.getY(), rect.getWidth(), rect.getHeight());
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
	
    @Override
    public String[] getResources() {
        return new String[]{"xlsx/Sport sign-up sheet.xlsx", "logo.png"};
    }
}
