package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.*;

import java.io.IOException;
import java.io.InputStream;

public class ConfigureDraft extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        //Set text.
        sheet.getRange("A1:G10").setValue("Text");

        //Add picture in sheet.
        InputStream stream = this.getResourceStream("logo.png");
        IShape picture = null;
        try {
            picture = sheet.getShapes().addPicture(stream, ImageType.PNG, 20, 20, 395, 60);
        } catch (IOException ioe) {

        }

        //Add header graphic.
        InputStream stream1 = this.getResourceStream("logo.png");
        sheet.getPageSetup().setCenterHeader("&G");
        sheet.getPageSetup().getCenterHeaderPicture().setGraphicStream(stream1, ImageType.PNG);
        sheet.getPageSetup().getCenterHeaderPicture().setWidth(100);
        sheet.getPageSetup().getCenterHeaderPicture().setHeight(13);

        //Set print without graphics in page content area.
        sheet.getPageSetup().setDraft(true);
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