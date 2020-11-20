package com.grapecity.documents.excel.examples.features.spreadjsfeatures;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.IBackgroundPicture;
import com.grapecity.documents.excel.drawing.ImageLayout;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.IOException;
import java.io.InputStream;

public class BackgroundImage extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

    	IWorksheet worksheet = workbook.getWorksheets().get(0);

        InputStream stream = this.getResourceStream("AcmeLogo.png");

        try {
			//Add background picture
			IBackgroundPicture picture = worksheet.getBackgroundPictures().addPictureInPixel(stream, ImageType.PNG, 10, 10, 500, 370);
			//Set image layout
			picture.setBackgroundImageLayout(ImageLayout.Zoom);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        //Set options
        workbook.getActiveSheet().getPageSetup().setPrintGridlines(true);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "AcmeLogo.png" };
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
	
		@Override
    public boolean getShowViewer() {
        return false;
    }
}
