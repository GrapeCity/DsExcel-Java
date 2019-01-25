package com.grapecity.documents.excel.examples.features.picture;

import java.io.IOException;
import java.io.InputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.IShape;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class PictureCopy extends ExampleBase {

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        InputStream stream = this.getResourceStream("logo.png");
        try {
            //Create a shape in worksheet, picture's range is Range("A2:I6")
            IShape picture = worksheet.getShapes().addPicture(stream, ImageType.PNG, 20, 20, 395, 60);
            //Range("A2:I6") must contain picture's range, copy a new picture to Range("J2:R6")
            worksheet.getRange("A2:I6").copy(worksheet.getRange("J2"));
            //worksheet.getRange("A2:I6").copy(worksheet.getRange("J2:R6"));

            //Cross sheet copy, copy a new shape to worksheet2's Range("J2:R6")
            //IWorksheet worksheet2 = workbook.getWorksheets().add();
            //worksheet.getRange("A2:I6").copy(worksheet2.getRange("J2"));
            //worksheet.getRange("A2:I6").copy(worksheet2.getRange("J2:R6"));
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    @Override
    public String[] getResources() {
        return new String[]{"logo.png"};
    }
}
