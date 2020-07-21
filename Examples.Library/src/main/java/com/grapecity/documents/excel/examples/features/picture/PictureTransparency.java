package com.grapecity.documents.excel.examples.features.picture;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.IOException;
import java.io.InputStream;

public class PictureTransparency extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {

        IWorksheet worksheet = workbook.getWorksheets().get(0);

        InputStream stream = this.getResourceStream("logo.png");
        try {
            IShape picture = worksheet.getShapes().addPicture(stream, ImageType.PNG, 20, 20, 790, 120);

            //Set picture's transparency as 60%.
            picture.getPictureFormat().setTransparency(0.6);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public String[] getResources() {
        return new String[]{"logo.png"};
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
    public boolean getIsNew() {
        return true;
    }
}
