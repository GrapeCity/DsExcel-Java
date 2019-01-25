package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.examples.*;

import java.io.IOException;
import java.io.InputStream;

public class ConfigureBlackAndWhite extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        InputStream stream = this.getResourceStream("logo.png");
        IShape picture = null;
        try {
            picture = sheet.getShapes().addPicture(stream, com.grapecity.documents.excel.drawing.ImageType.PNG, 20, 20, 395, 60);
        } catch (IOException ioe) {

        }

        //Set text font color.
        sheet.getRange("A1:D4").setValue("Font");
        sheet.getRange("A1:D4").getFont().setColor(Color.GetRed());

        //Set cell color
        sheet.getRange("A7:D10").setValue("Green");
        sheet.getRange("A7:D10").getInterior().setColor(Color.GetGreen());

        //Set print black and white.
        sheet.getPageSetup().setBlackAndWhite(true);
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
        return new String[]{"logo.png"};
    }

}