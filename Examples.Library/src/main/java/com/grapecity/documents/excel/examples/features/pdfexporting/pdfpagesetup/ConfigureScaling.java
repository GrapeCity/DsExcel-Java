package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.*;

import java.io.IOException;
import java.io.InputStream;

public class ConfigureScaling extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        InputStream stream = this.getResourceStream("logo.png");
        IShape picture = null;
        try {
            picture = sheet.getShapes().addPicture(stream, ImageType.PNG, 20, 20, 395, 60);
        } catch (IOException ioe) {

        }
        sheet.getRange("B2:D4").setValue("Text");

        sheet.getPageSetup().setPrintGridlines(true);

        //Set scaling.
        sheet.getPageSetup().setZoom(200);
    }

    @Override
    public boolean getSavePdf() {
        return true;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
}