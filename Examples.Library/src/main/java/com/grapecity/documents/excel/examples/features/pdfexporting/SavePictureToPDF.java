package com.grapecity.documents.excel.examples.features.pdfexporting;

import java.io.IOException;
import java.io.InputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PageOrientation;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SavePictureToPDF extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getPageSetup().setOrientation(PageOrientation.Landscape);
        InputStream stream = this.getResourceStream("logo.png");
        IShape picture = null;
        try {
            picture = worksheet.getShapes().addPicture(stream, ImageType.PNG, 20, 20, 400, 100);
        } catch (IOException ioe) {

        }
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