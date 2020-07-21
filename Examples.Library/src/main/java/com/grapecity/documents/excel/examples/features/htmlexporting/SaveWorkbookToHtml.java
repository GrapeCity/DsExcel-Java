package com.grapecity.documents.excel.examples.features.htmlexporting;

import com.grapecity.documents.excel.SaveFileFormat;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.examples.ExampleBase;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

public class SaveWorkbookToHtml extends ExampleBase {
    @Override
    public void execute(Workbook workbook, ByteArrayOutputStream outputStream) {
        // Open an excel file
        InputStream fileStream = this.getResourceStream("xlsx/Project tracker.xlsx");
        workbook.open(fileStream);

        workbook.save(outputStream, SaveFileFormat.Html);
    }

    @Override
    public boolean getSaveAsZip() {
        return true;
    }

    @Override
    public boolean getIsNew() {
        return true;
    }

    @Override
    public String[] getResources() {
        return new String[] { "xlsx/Project tracker.xlsx" };
    }
}
