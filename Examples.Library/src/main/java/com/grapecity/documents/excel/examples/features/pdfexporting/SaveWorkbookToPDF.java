package com.grapecity.documents.excel.examples.features.pdfexporting;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.SaveFileFormat;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

public class SaveWorkbookToPDF extends ExampleBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet sheet1 = workbook.getWorksheets().get(0);

        //set style.
        sheet1.getRange("A1").setValue("Sheet1");
        sheet1.getRange("A1").getFont().setName("Wide Latin");
        sheet1.getRange("A1").getFont().setColor(Color.GetRed());
        sheet1.getRange("A1").getInterior().setColor(Color.GetGreen());
        
        
        IWorksheet sheet2 = workbook.getWorksheets().add();
        
        //set style.
        sheet2.getRange("A1").setValue("Sheet2");
        sheet2.getRange("A1").getFont().setName("Wide Latin");
        sheet2.getRange("A1").getFont().setColor(Color.GetRed());
        sheet2.getRange("A1").getInterior().setColor(Color.GetYellow());

        //change the path to real export path when save.
        workbook.save("dest.pdf", SaveFileFormat.Pdf);
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