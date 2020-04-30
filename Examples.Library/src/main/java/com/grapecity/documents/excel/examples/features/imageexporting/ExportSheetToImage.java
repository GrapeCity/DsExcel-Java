package com.grapecity.documents.excel.examples.features.imageexporting;

import java.io.ByteArrayOutputStream;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

public class ExportSheetToImage extends ExampleBase {
    @Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        
        workbook.open(this.getResourceStream("xlsx/Home inventory.xlsx"));
        
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Save the worksheet as image to a stream.
        worksheet.toImage(outputStream, ImageType.PNG);
    }

    @Override
    public String getTemplateName() {
        return "Home inventory.xlsx";
    }

    @Override
    public boolean getSaveAsImage() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] {"xlsx/Home inventory.xlsx"};
    }

    @Override
    public boolean getIsNew() {
        return true;
    }
}