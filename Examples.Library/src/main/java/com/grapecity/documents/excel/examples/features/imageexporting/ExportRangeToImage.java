package com.grapecity.documents.excel.examples.features.imageexporting;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.ImageType;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

public class ExportRangeToImage extends ExampleBase {
    @Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/Personal Monthly Budget.xlsx");
        workbook.open(fileStream);
        IWorksheet worksheet = workbook.getWorksheets().get(0);

        // Export range "B14:E25" to image
        worksheet.getRange("B14:E25").toImage(outputStream, ImageType.PNG);
    }

    @Override
    public boolean getSaveAsImage() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/Personal Monthly Budget.xlsx" };
    }
}
