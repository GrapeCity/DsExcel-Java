package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.io.IOException;
import java.io.InputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.IBackgroundPicture;
import com.grapecity.documents.excel.drawing.ImageLayout;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SaveBackgroundPicturesToPDF extends ExampleBase {
	@Override
	public void execute(Workbook workbook) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/To_Do_List.xlsx");
        workbook.open(fileStream);
        
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Set a background image for worksheet
        InputStream stream = this.getResourceStream("AcmeLogo.png");

        // Add a background picture for the worksheet, and the background picture will be rendered into the destination rectangle[10, 10, 500, 370].
        IBackgroundPicture picture = null;
        try {
            picture = worksheet.getBackgroundPictures().addPictureInPixel(stream, ImageType.PNG, 10, 10, 500, 370);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        // The background picture will be resized to fill the destination dimensions.
        picture.setBackgroundImageLayout(ImageLayout.Tile);

        // Sets the transparency of the background pictures.
        picture.setTransparency(0.5);
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
	public String[] getResources() {
        return new String[] { "xlsx/To_Do_List.xlsx", "AcmeLogo.png" };
	}
}
