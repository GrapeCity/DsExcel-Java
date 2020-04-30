package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.examples.*;

import java.util.*;

public class SetFontsFolderPath extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet = workbook.getWorksheets().get(0);

        //set style.
        sheet.getRange("A1").setValue("Sheet1");
        sheet.getRange("A1").getFont().setName("Wide Latin");
        sheet.getRange("A1").getFont().setColor(Color.GetRed());
        sheet.getRange("A1").getInterior().setColor(Color.GetGreen());

        //specify font path.
        Workbook.FontsFolderPath = "/var/usr/public/Fonts";

//        //get the used fonts list in workbook, the list are:"Wide Latin", "Calibri"
        List<FontInfo> fonts = workbook.getUsedFonts();

        //change the path to real export path when save.
        sheet.save("dest.pdf", SaveFileFormat.Pdf);
    }

    @Override
    public boolean getCanDownload() {
        return false;
    }

    @Override
    public boolean getShowViewer() {
        return false;
    }
}