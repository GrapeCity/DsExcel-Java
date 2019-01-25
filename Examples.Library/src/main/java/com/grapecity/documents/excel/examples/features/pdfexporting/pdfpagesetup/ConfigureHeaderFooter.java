package com.grapecity.documents.excel.examples.features.pdfexporting.pdfpagesetup;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.*;

import java.io.InputStream;

public class ConfigureHeaderFooter extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        //Set data.
        sheet.getRange("A1:G10").setValue("Text");

        //Set page header.
        sheet.getPageSetup().setLeftHeader("&\"Arial,Italic\"LeftHeader");
        sheet.getPageSetup().setRightHeader("&KFF0000com.grapecity");
        sheet.getPageSetup().setCenterHeader("&P");

        //Set page footer picture.
        InputStream stream = this.getResourceStream("logo.png");
        sheet.getPageSetup().setCenterFooter("&G");
        sheet.getPageSetup().getCenterFooterPicture().setGraphicStream(stream, ImageType.PNG);
        sheet.getPageSetup().getCenterFooterPicture().setWidth(100);
        sheet.getPageSetup().getCenterFooterPicture().setHeight(13);
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